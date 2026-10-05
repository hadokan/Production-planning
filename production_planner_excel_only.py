#!/usr/bin/env python3
"""Excel v6 -> constructive search + unrestricted MILP -> five-sheet Excel.

No ChatGPT/API connection or commercial solver is required. All computation is
local. Read README_TR.md for supported schema, numerical conventions and limits.
"""
from __future__ import annotations

import argparse
import hashlib
import logging
import math
import os
import platform
import re
import sys
import time
import warnings
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import numpy as np
import scipy
from scipy.optimize import Bounds, LinearConstraint, milp
from scipy.sparse import coo_matrix
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

try:
    from numba import njit
    NUMBA_AVAILABLE = True
except ImportError:
    NUMBA_AVAILABLE = False
    def njit(*args, **kwargs):
        def wrapper(f):
            return f
        return wrapper

VERSION = "6.2.0"
TOL = 1e-5                 # Validation tolerance in tons / hours.
POSITIVE_T = 1e-4          # Numerical activation floor: 0.1 kg, NOT a batch size.
LOG = logging.getLogger("production_planner")

SCHEMA = {
    "products": "Product_ID Product Family Salt_Status Sequence_Order Priority".split(),
    "machines": "Machine_ID Machine Active Priority".split(),
    "machine_product": "Machine_ID Product_ID Allowed MachinePriority Capacity_tph MinRun_t MaxDaily_t Preparation_min".split(),
    "targets": "Product_ID Product Target_t TargetPriority TargetMode".split(),
    "calendar": "Machine_ID Day Shift Available Hours Reason".split(),
    "sequence_rules": "Rule_ID Family Before_Product_ID After_Product_ID Scope Hard".split(),
    "changeover_rules": "Rule_ID Scope Max_Products_Per_Shift Max_Changeovers_Per_Day Hard Description".split(),
    "allocation_rules": "Rule_ID Scope Prefer_Single_Machine Preferred_Max_Machines Allow_Split_If_Required Hard Description Objective_ID".split(),
    "initial_state": "Machine_ID Current_Product_ID Run_Produced_t Setup_Completed".split(),
    "objectives": "Priority Objective_ID Direction Enabled Description".split(),
    "settings": "Parameter Value Description".split(),
}
SETTING_NAMES = set("""SchemaVersion ShiftHours DefaultPreparationMin MaxPlanningDays
AllowFinalRunBelowMinimum AllowSurplusProduction PreventProductReturn
PreferContinuousProduction CriticalProductFirst ObjectiveMode UnmetTolerance_t
UnmetToleranceScope UnmetReferenceMode CountCrossDayChangeovers
ChangeoverAttribution CountInitialSetupAsChangeover IdleResetsProductState
SequenceRuleSource SequenceConflictPolicy OptimizationMode HeuristicTrials
MILPPolishTimeSec MILPNeighborhoodMode RuntimeLogging InputValidation
UnknownParameterPolicy RandomSeed""".split())
OBJECTIVE_NAMES = ["MIN_UNMET", "MIN_PRODUCT_MACHINE_SPLITS", "MIN_CHANGEOVERS",
    "MIN_FAMILY_MACHINE_SPLITS", "MIN_PRODUCTS_PER_MACHINE", "MIN_SURPLUS",
    "MIN_MACHINE_PRIORITY_COST"]

class InputError(ValueError):
    """Input is missing, inconsistent, or outside the implemented v6 contract."""


def text(v: Any) -> str:
    return " ".join(str(v).strip().split()) if v is not None else ""


def number(v: Any, name: str, minimum: float = 0, integer: bool = False) -> float:
    if isinstance(v, bool):
        raise InputError(f"{name}: use numeric 0/1, not a Boolean cell.")
    try:
        f = float(v)
    except (ValueError, TypeError) as exc:
        raise InputError(f"{name}: expected a number; got {v!r}") from exc
    if not math.isfinite(f) or f < minimum or (integer and f != int(f)):
        raise InputError(f"{name}: invalid numeric value {v!r}")
    return int(f) if integer else f


def flag(v: Any, name: str) -> int:
    k = number(v, name, integer=True)
    if k not in (0, 1):
        raise InputError(f"{name}: expected 0 or 1.")
    return k


def read_tables(path: Path) -> dict[str, list[dict]]:
    wb = load_workbook(path, data_only=False, read_only=True)
    try:
        if set(wb.sheetnames) != set(SCHEMA):
            raise InputError(f"Sheet mismatch. Missing={set(SCHEMA)-set(wb.sheetnames)}; "
                             f"unknown={set(wb.sheetnames)-set(SCHEMA)}")
        tables = {}
        for name, expected in SCHEMA.items():
            rows = list(wb[name].iter_rows(values_only=True))
            if not rows:
                raise InputError(f"{name}: empty sheet.")
            raw = list(rows[0])
            while raw and raw[-1] is None:
                raw.pop()
            headers = [text(v) for v in raw]
            if len(headers) != len(set(headers)) or set(headers) != set(expected):
                raise InputError(f"{name}: expected headers {expected}; got {headers}")
            data = []
            for rn, vals in enumerate(rows[1:], 2):
                if not any(v is not None for v in vals):
                    continue
                if any(v is not None for v in vals[len(headers):]):
                    raise InputError(f"{name}!row {rn}: data beyond known columns.")
                d = dict(zip(headers, vals))
                if any(isinstance(v, str) and v.startswith("=") for v in d.values()):
                    raise InputError(f"{name}!row {rn}: input cells must be values, not formulas.")
                d["_row"] = rn
                data.append(d)
            tables[name] = data
        return tables
    finally:
        wb.close()


def unique(rows: list[dict], keys: tuple[str, ...], sheet: str) -> None:
    seen = set()
    for r in rows:
        key = tuple(text(r[k]) for k in keys)
        if any(not x for x in key) or key in seen:
            raise InputError(f"{sheet}: empty or duplicate key {key}")
        seen.add(key)


@dataclass
class Data:
    path: Path
    sha256: str
    tables: dict
    settings: dict
    products: list[dict]
    machines: list[dict]
    slots: list[tuple[int, int]]
    families: list[str]
    family_index: np.ndarray
    rate: np.ndarray
    minimum: np.ndarray
    daily: np.ndarray
    prep: np.ndarray
    priority: np.ndarray
    hours: np.ndarray
    target: np.ndarray
    hard: np.ndarray
    predecessors: np.ndarray
    initial: np.ndarray
    initial_qty: np.ndarray
    prepared: np.ndarray
    max_changes: np.ndarray
    product_limits: np.ndarray
    family_limits: np.ndarray
    objectives: list[str]
    warnings: list[str] = field(default_factory=list)

    @property
    def M(self): return len(self.machines)
    @property
    def P(self): return len(self.products)
    @property
    def T(self): return len(self.slots)


def load_input(path: str | Path) -> Data:
    path = Path(path).expanduser().resolve()
    if not path.is_file() or path.suffix.lower() != ".xlsx":
        raise InputError(f"Input .xlsx not found: {path}")
    tb = read_tables(path)
    for sh, keys in {"products": ("Product_ID",), "machines": ("Machine_ID",),
         "machine_product": ("Machine_ID", "Product_ID"), "targets": ("Product_ID",),
         "calendar": ("Machine_ID", "Day", "Shift"), "sequence_rules": ("Rule_ID",),
         "changeover_rules": ("Rule_ID",), "allocation_rules": ("Rule_ID",),
         "initial_state": ("Machine_ID",), "objectives": ("Priority",),
         "settings": ("Parameter",)}.items():
        unique(tb[sh], keys, sh)
    s = {text(r['Parameter']): r['Value'] for r in tb['settings']}
    if set(s) != SETTING_NAMES:
        raise InputError(f"Settings mismatch. Missing={SETTING_NAMES-set(s)}; unknown={set(s)-SETTING_NAMES}")
    required_values = {
        "SchemaVersion": "6.0", "ObjectiveMode": "LEXICOGRAPHIC_WITH_UNMET_TOLERANCE",
        "UnmetToleranceScope": "TOTAL_PLAN", "UnmetReferenceMode": "BEST_FOUND_BASELINE",
        "ChangeoverAttribution": "NEW_PRODUCT_START_DAY", "SequenceRuleSource": "sequence_rules",
        "SequenceConflictPolicy": "ERROR", "OptimizationMode": "HYBRID",
        "MILPNeighborhoodMode": "REALLOCATION_AND_RESEQUENCING", "InputValidation": "STRICT",
        "UnknownParameterPolicy": "ERROR"}
    for key, wanted in required_values.items():
        if text(s[key]) != wanted:
            raise InputError(f"Unsupported {key}={s[key]!r}. This release supports {wanted!r}.")
    for key in ['AllowFinalRunBelowMinimum','PreferContinuousProduction','CriticalProductFirst','RuntimeLogging']:
        s[key] = flag(s[key], key)
    for key, wanted in {"AllowSurplusProduction":0,"PreventProductReturn":1,
             "CountCrossDayChangeovers":1,"CountInitialSetupAsChangeover":0,"IdleResetsProductState":0}.items():
        if flag(s[key], key) != wanted:
            raise InputError(f"Unsupported {key}={s[key]}; this release requires {wanted}.")
    for key in ['DefaultPreparationMin','MILPPolishTimeSec','UnmetTolerance_t']:
        s[key] = number(s[key], key)
    s['ShiftHours'] = number(s['ShiftHours'], 'ShiftHours', minimum=1e-9)
    for key in ['MaxPlanningDays','HeuristicTrials']:
        s[key] = number(s[key], key, minimum=1, integer=True)
    s['RandomSeed'] = number(s['RandomSeed'], 'RandomSeed', minimum=0, integer=True)
    if s['RandomSeed'] > 2147483646:
        raise InputError('RandomSeed must be between 0 and 2147483646.')
    products = tb['products']
    if s['RuntimeLogging'] != 1:
        raise InputError('This audited release requires RuntimeLogging=1.')
    if not products or not tb['machines']:
        raise InputError('At least one product and one machine are required.')
    for r in products:
        for k in ['Product_ID','Product','Family','Salt_Status']:
            r[k] = text(r[k])
            if not r[k]: raise InputError(f"products: {k} is empty.")
        if r['Salt_Status'] not in ['UNSALTED','SALTED','NONE']:
            raise InputError(f"Invalid Salt_Status for {r['Product_ID']}")
        r['Sequence_Order'] = number(r['Sequence_Order'],'Sequence_Order',minimum=1,integer=True)
        r['Priority'] = number(r['Priority'],'Product.Priority',minimum=1,integer=True)
    if len({r['Priority'] for r in products}) > 1:
        raise InputError('Mixed product Priority tiers are not implemented; equal values are required in this v6 release.')
    allmachines = {text(r['Machine_ID']): r for r in tb['machines']}
    machines = []
    for r in tb['machines']:
        r['Machine_ID'],r['Machine'] = text(r['Machine_ID']),text(r['Machine'])
        r['Priority'] = number(r['Priority'],'Machine.Priority',minimum=1,integer=True)
        r['Active'] = flag(r['Active'],'Active')
        if r['Active']: machines.append(r)
    if not machines: raise InputError('No active machines.')
    pm = {r['Product_ID']: i for i,r in enumerate(products)}
    mm = {r['Machine_ID']: i for i,r in enumerate(machines)}
    M,P = len(machines),len(products)
    arrays = [np.zeros((M,P),dtype=float) for _ in range(5)]
    rate,minimum,daily,prep,priority = arrays
    for r in tb['machine_product']:
        mid,pid = text(r['Machine_ID']),text(r['Product_ID'])
        if mid not in allmachines or pid not in pm: raise InputError(f"Unknown machine-product: {mid}/{pid}")
        allowed = flag(r['Allowed'],'Allowed')
        vals = [number(r[k],f'{mid}/{pid}/{k}') for k in ['Capacity_tph','MinRun_t','MaxDaily_t']]
        pr = number(r['Preparation_min'] if r['Preparation_min'] is not None else s['DefaultPreparationMin'],'Preparation_min')
        mp = number(r['MachinePriority'],'MachinePriority',minimum=1,integer=True) if allowed else 0
        if allowed and (vals[0] <= 0 or vals[2] <= 0):
            raise InputError(f'{mid}/{pid}: allowed pair needs positive rate and MaxDaily_t.')
        if not allowed and any(vals): raise InputError(f'{mid}/{pid}: disallowed pair must have zero rate/run/daily quantities.')
        if mid in mm and allowed:
            m,p=mm[mid],pm[pid]
            rate[m,p],minimum[m,p],daily[m,p],prep[m,p],priority[m,p]=*vals,pr,mp
    target,hard=np.zeros(P),np.zeros(P,dtype=np.int64)
    if {text(r['Product_ID']) for r in tb['targets']} != set(pm):
        raise InputError('targets must contain exactly one row for each product (zero target is allowed).')
    target_priorities=[]
    for r in tb['targets']:
        p=pm[text(r['Product_ID'])]
        if text(r['Product']) != products[p]['Product']:
            raise InputError(f"targets.Product conflicts with products for {r['Product_ID']}")
        target[p]=number(r['Target_t'],'Target_t')
        if 0<target[p]<POSITIVE_T: raise InputError('Target below numerical activation floor of 0.0001 t.')
        mode=text(r['TargetMode'])
        if mode not in ['SOFT','HARD']: raise InputError('TargetMode must be SOFT or HARD.')
        hard[p]=(mode=='HARD')
        target_priorities.append(number(r['TargetPriority'],'TargetPriority',minimum=1,integer=True))
    if len(set(target_priorities))>1:
        raise InputError('This v6 release supports equal TargetPriority values. Mixed target tiers need an explicitly agreed tolerance policy; not silently weighted.')
    slots_set=set(); cal={}
    for r in tb['calendar']:
        mid=text(r['Machine_ID'])
        match=re.fullmatch(r'Day-(\d+)',text(r['Day']))
        if mid not in allmachines or not match: raise InputError('calendar: unknown machine or invalid Day-n label.')
        day=int(match.group(1)); shift=number(r['Shift'],'Shift',minimum=1,integer=True)
        if not 1<=day<=s['MaxPlanningDays']: raise InputError('Calendar exceeds MaxPlanningDays.')
        available=flag(r['Available'],'Available'); hours=number(r['Hours'],'Hours')
        if (not available and hours!=0) or (available and hours<=0): raise InputError('Calendar Available/Hours conflict.')
        if mid in mm:
            key=(mm[mid],day,shift)
            if key in cal: raise InputError('Duplicate normalized calendar key.')
            cal[key]=hours
            slots_set.add((day,shift))
    slots=sorted(slots_set); T=len(slots)
    if not T: raise InputError('Empty active calendar.')
    hours=np.zeros((M,T))
    for m in range(M):
        for t,(day,shift) in enumerate(slots):
            if (m,day,shift) not in cal: raise InputError(f'Missing calendar row for {machines[m]["Machine_ID"]}/Day-{day}/{shift}; add explicit closure.')
            hours[m,t]=cal[m,day,shift]
    pre=np.zeros((P,P),dtype=np.int64)
    for r in tb['sequence_rules']:
        a,b=text(r['Before_Product_ID']),text(r['After_Product_ID'])
        if a not in pm or b not in pm or a==b: raise InputError('Invalid sequence pair.')
        if text(r['Scope'])!='SAME_MACHINE' or flag(r['Hard'],'Sequence.Hard')!=1:
            raise InputError('Only HARD SAME_MACHINE sequence rules are supported.')
        pa,pb=products[pm[a]],products[pm[b]]
        if pa['Family']!=pb['Family'] or text(r['Family'])!=pa['Family'] or pa['Sequence_Order']>=pb['Sequence_Order']:
            raise InputError(f'Sequence rule conflicts with Family/Sequence_Order: {a}->{b}')
        if pre[pm[a],pm[b]]: raise InputError('Duplicate sequence pair.')
        pre[pm[a],pm[b]]=1
    for a in range(P):
        for b in range(P):
            if products[a]['Family']==products[b]['Family'] and products[a]['Sequence_Order']<products[b]['Sequence_Order'] and not pre[a,b]:
                raise InputError(f'Missing explicit sequence pair {products[a]["Product_ID"]}->{products[b]["Product_ID"]}')
    initial=np.full(M,-1,dtype=np.int64); initial_qty=np.zeros(M); prepared=np.zeros(M,dtype=np.int64)
    if {text(r['Machine_ID']) for r in tb['initial_state']}!=set(allmachines):
        raise InputError('initial_state needs one row for each machine.')
    for r in tb['initial_state']:
        mid=text(r['Machine_ID']); pid=text(r['Current_Product_ID'])
        if pid and pid not in pm: raise InputError('Unknown initial product.')
        q=number(r['Run_Produced_t'],'Run_Produced_t'); ready=flag(r['Setup_Completed'],'Setup_Completed')
        if not pid and (q or ready): raise InputError('Empty initial product requires zero Run_Produced_t and Setup_Completed.')
        if mid in mm:
            m=mm[mid]
            if pid:
                p=pm[pid]
                if rate[m,p]<=0: raise InputError('Initial product is not allowed on its machine.')
                # Abandoning an incomplete carried run requires a separate terminal-state contract.
                if q>0 and q+TOL<minimum[m,p]:
                    raise InputError(f'{mid}: incomplete carried run is not supported in this release. Do not drop its balance silently; model a locked continuation first.')
                initial[m]=p
            initial_qty[m]=q; prepared[m]=ready
    cr=tb['changeover_rules']
    if len(cr)!=1 or text(cr[0]['Scope'])!='MACHINE_DAY': raise InputError('Use one global MACHINE_DAY changeover rule.')
    if number(cr[0]['Max_Products_Per_Shift'],'Max_Products_Per_Shift',integer=True)!=1 or flag(cr[0]['Hard'],'Changeover.Hard')!=1:
        raise InputError('This release requires HARD one-product-per-machine-shift.')
    maxchg=number(cr[0]['Max_Changeovers_Per_Day'],'Max_Changeovers_Per_Day',integer=True)
    max_changes=np.full(M,maxchg,dtype=np.int64)
    families=list(dict.fromkeys(r['Family'] for r in products))
    family_index=np.array([families.index(r['Family']) for r in products],dtype=np.int64)
    product_limits=np.full(P,M,dtype=np.int64); family_limits=np.full(len(families),M,dtype=np.int64)
    for r in tb['allocation_rules']:
        scope=text(r['Scope']); oid=text(r['Objective_ID'])
        preferred=number(r['Preferred_Max_Machines'],'Preferred_Max_Machines',minimum=1,integer=True)
        prefer=flag(r['Prefer_Single_Machine'],'Prefer_Single_Machine')
        split=flag(r['Allow_Split_If_Required'],'Allow_Split_If_Required'); ishard=flag(r['Hard'],'Allocation.Hard')
        if not prefer or preferred!=1: raise InputError('Only a one-machine preference/limit is implemented; unsupported allocation values are rejected.')
        if scope=='ALL_PRODUCTS': ids=list(range(P)); dest=product_limits; expected=OBJECTIVE_NAMES[1]
        elif scope.startswith('PRODUCT:') and scope[8:] in pm: ids=[pm[scope[8:]]]; dest=product_limits; expected=OBJECTIVE_NAMES[1]
        elif scope=='ALL_FAMILIES': ids=list(range(len(families))); dest=family_limits; expected=OBJECTIVE_NAMES[3]
        elif scope.startswith('FAMILY:') and scope[7:] in families: ids=[families.index(scope[7:])]; dest=family_limits; expected=OBJECTIVE_NAMES[3]
        else: raise InputError(f'Unsupported allocation scope {scope}')
        if oid!=expected: raise InputError('Allocation Objective_ID mismatch.')
        if ishard or not split:
            for i in ids: dest[i]=min(dest[i],preferred)
    objs=[]; seen=set()
    for r in sorted(tb['objectives'],key=lambda z:number(z['Priority'],'Objective.Priority',minimum=1,integer=True)):
        oid=text(r['Objective_ID'])
        if oid not in OBJECTIVE_NAMES or oid in seen or text(r['Direction'])!='MIN': raise InputError('Unknown/duplicate objective or unsupported direction.')
        seen.add(oid)
        if flag(r['Enabled'],'Enabled'): objs.append(oid)
    if not objs or objs[0]!='MIN_UNMET': raise InputError('MIN_UNMET must be the first enabled objective.')
    notes=[]
    for p in range(P):
        if target[p]>0 and not np.any(rate[:,p]>0):
            notes.append(f"NO_ELIGIBLE_MACHINE: {products[p]['Product_ID']}")
    return Data(path,hashlib.sha256(path.read_bytes()).hexdigest(),tb,s,products,machines,slots,
        families,family_index,rate,minimum,daily,prep,priority,hours,target,hard,pre,
        initial,initial_qty,prepared,max_changes,product_limits,family_limits,objs,notes)


@dataclass
class Plan:
    product: np.ndarray
    quantity: np.ndarray
    method: str = ''
    metrics: dict = field(default_factory=dict)

    def clone(self):
        return Plan(self.product.copy(),self.quantity.copy(),self.method,dict(self.metrics))


def evaluate(d: Data, plan: Plan) -> dict:
    amounts=np.zeros(d.P); assigned=np.zeros((d.M,d.P),dtype=int)
    changes=0; run_count=0; pcost=0.; mastercost=0.; idle_gaps=0
    for m in range(d.M):
        last=int(d.initial[m]); seen=set(); had=False; idle_since=False
        for t in range(d.T):
            p=int(plan.product[m,t]); q=float(plan.quantity[m,t])
            if p<0 or q<=TOL:
                if had and d.hours[m,t]>0: idle_since=True
                continue
            if idle_since and p==last: idle_gaps+=1
            idle_since=False; had=True
            amounts[p]+=q; assigned[m,p]=1
            pcost+=q*(d.priority[m,p]-1)
            mastercost+=q*(d.machines[m]['Priority']-1)
            if p not in seen: run_count+=1; seen.add(p)
            if last>=0 and p!=last: changes+=1
            last=p
    uses=assigned.sum(axis=0); famuses=[]
    for f in range(len(d.families)):
        famuses.append(int(np.any(assigned[:,d.family_index==f],axis=1).sum()))
    result={
        'target':float(d.target.sum()),'planned':float(amounts.sum()),
        'unmet':float(np.maximum(d.target-amounts,0).sum()),
        'surplus':float(np.maximum(amounts-d.target,0).sum()),
        'product_splits':int(np.maximum(uses-1,0).sum()),
        'split_products':int((uses>1).sum()),'changeovers':changes,
        'family_splits':sum(max(x-1,0) for x in famuses),
        'assignments':int(assigned.sum()),'runs':run_count,
        'machines_used':int(np.any(assigned,axis=1).sum()),
        'priority_cost':float(pcost),'master_priority_cost':float(mastercost),'idle_gaps':idle_gaps,
    }
    plan.metrics=result
    return result


def preference_key(d: Data, met: dict) -> tuple:
    lookup={'MIN_UNMET':'unmet','MIN_PRODUCT_MACHINE_SPLITS':'product_splits',
        'MIN_CHANGEOVERS':'changeovers','MIN_FAMILY_MACHINE_SPLITS':'family_splits',
        'MIN_PRODUCTS_PER_MACHINE':'assignments','MIN_SURPLUS':'surplus',
        'MIN_MACHINE_PRIORITY_COST':'priority_cost'}
    k=tuple(round(met[lookup[o]],6) for o in d.objectives if o!='MIN_UNMET')
    return k+(round(met['unmet'],6),round(met['master_priority_cost'],6),
              met['idle_gaps'] if d.settings['PreferContinuousProduction'] else 0)


def validate(d: Data, plan: Plan, unmet_cap: float | None = None) -> list[dict]:
    """Recompute constraints from raw assignments. Never trust solver flags."""
    bad=[]
    def fail(code, message): bad.append({'code':code,'message':message})
    if plan.product.shape!=(d.M,d.T) or plan.quantity.shape!=(d.M,d.T):
        return [{'code':'SHAPE','message':'Plan dimensions do not match input.'}]
    if not np.isfinite(plan.quantity).all(): return [{'code':'NUMBER','message':'Non-finite quantities.'}]
    amounts=np.zeros(d.P); dayqty=defaultdict(float); changes=defaultdict(int)
    runs=defaultdict(list); assigned=np.zeros((d.M,d.P),dtype=int)
    for m in range(d.M):
        last=int(d.initial[m]); seen={last} if last>=0 else set(); first_prod=True
        for t,(day,shift) in enumerate(d.slots):
            p=int(plan.product[m,t]); q=float(plan.quantity[m,t]); where=f'{d.machines[m]["Machine"]}/Day-{day}/S{shift}'
            if p<0:
                if p != -1: fail('PRODUCT_ID',where)
                if abs(q)>TOL: fail('EMPTY_QTY',where)
                continue
            if p>=d.P or q<POSITIVE_T-TOL:
                fail('ACTIVE_QTY',where); continue
            if d.hours[m,t]<=0: fail('CALENDAR',where)
            if d.rate[m,p]<=0: fail('ELIGIBILITY',where); continue
            new_run=p!=last or first_prod
            setup=d.prep[m,p] if new_run else 0.
            if first_prod and p==d.initial[m] and d.prepared[m]: setup=0.
            used=q/d.rate[m,p]+setup/60.
            if used>d.hours[m,t]+TOL: fail('SHIFT_TIME',f'{where}: {used:.8f} > {d.hours[m,t]} h')
            if p!=last:
                if p in seen: fail('NO_RETURN',where)
                for previous in seen:
                    if d.predecessors[p,previous]: fail('SEQUENCE',f'{where}: {p} must precede {previous}')
                if last>=0: changes[m,day]+=1
            seen.add(p); last=p; first_prod=False
            amounts[p]+=q; dayqty[m,p,day]+=q; assigned[m,p]=1
            runs[m,p].append((t,q))
    for (m,p,day),q in dayqty.items():
        if q>d.daily[m,p]+TOL: fail('DAILY_TON',f'{d.machines[m]["Machine"]}/{d.products[p]["Product_ID"]}/Day-{day}: {q:.6f}>{d.daily[m,p]}')
    for (m,day),n in changes.items():
        if n>d.max_changes[m]: fail('DAILY_CHANGEOVER',f'{d.machines[m]["Machine"]}/Day-{day}: {n}>{d.max_changes[m]}')
    for p in range(d.P):
        if amounts[p]>d.target[p]+TOL: fail('SURPLUS',d.products[p]['Product_ID'])
        if d.hard[p] and amounts[p]<d.target[p]-TOL: fail('HARD_TARGET',d.products[p]['Product_ID'])
        if assigned[:,p].sum()>d.product_limits[p]: fail('ALLOCATION_PRODUCT',d.products[p]['Product_ID'])
    for f in range(len(d.families)):
        if np.any(assigned[:,d.family_index==f],axis=1).sum()>d.family_limits[f]: fail('ALLOCATION_FAMILY',d.families[f])
    for (m,p),rows in runs.items():
        q=sum(v for _,v in rows)
        credit=d.initial_qty[m] if d.initial[m]==p else 0.
        if q+credit+TOL>=d.minimum[m,p]: continue
        # All other same-product runs must have ended before this final run starts.
        other=[items for (other_m,other_p),items in runs.items() if other_p==p and other_m!=m]
        temporal=all(max(t for t,_ in items)<rows[0][0] for items in other)
        if not (d.settings['AllowFinalRunBelowMinimum'] and abs(amounts[p]-d.target[p])<=TOL and temporal):
            fail('MIN_RUN',f'{d.machines[m]["Machine"]}/{d.products[p]["Product_ID"]}: {q:.6f}+{credit:.6f} < {d.minimum[m,p]}; not an authorized final remainder')
    met=evaluate(d,plan)
    if unmet_cap is not None and met['unmet']>unmet_cap+TOL: fail('UNMET_CAP',f'{met["unmet"]} > {unmet_cap}')
    return bad


@njit(cache=True)
def _construct_batch(count, seed, rate, minimum, daily, prep, hours, target,
                     pre, init, init_qty, prepared, maxchg, days, fam, priority,
                     allow_final, critical, singlelimits, familylimits):
    """Randomized block construction; no product/machine names are hardcoded."""
    np.random.seed(seed)
    M,P=rate.shape; T=hours.shape[1]
    X=np.full((count,M,T),-1,dtype=np.int16); Q=np.zeros((count,M,T))
    for n in range(count):
        rem=target.copy(); lastend=np.full(P,-1,dtype=np.int64)
        used=np.zeros((M,P),dtype=np.int64)
        morder=np.argsort(np.random.random(M))
        if critical and np.random.random()<0.7:
            # Restricted machines first, random ties.
            keys=np.random.random(M)
            for m in range(M):
                keys[m]+=np.sum(rate[m,:]>0)
            morder=np.argsort(keys)
        for oi in range(M):
            m=morder[oi]; cursor=0; last=init[m]; first=True
            done=np.zeros(P,dtype=np.int64)
            day_changes=np.zeros(np.max(days)+1,dtype=np.int64)
            while cursor<T:
                porder=np.argsort(np.random.random(P))
                chosen=False
                for pi in range(P):
                    p=porder[pi]
                    if rem[p]<POSITIVE_T or rate[m,p]<=0 or done[p]: continue
                    if p==init[m] and not first and last!=p: continue
                    blocked=False
                    for r in range(P):
                        if pre[p,r] and (done[r] or init[m]==r): blocked=True
                    if blocked: continue
                    if np.sum(used[:,p])>=singlelimits[p]: continue
                    f=fam[p]; fammachines=0; already=False
                    for mm in range(M):
                        has=False
                        for pp in range(P):
                            if fam[pp]==f and used[mm,pp]: has=True
                        if has: fammachines+=1
                        if mm==m: already=has
                    if not already and fammachines>=familylimits[f]: continue
                    credit=init_qty[m] if first and init[m]==p else 0.
                    need=max(0.,minimum[m,p]-credit)
                    st=cursor
                    # Target-limited last remainder must follow all other runs of this product.
                    if rem[p]+TOL<need:
                        if not allow_final: continue
                        st=max(st,lastend[p]+1)
                    setup=0. if first and init[m]==p and prepared[m] else prep[m,p]
                    ischange=(last>=0 and p!=last)
                    while st<T:
                        if hours[m,st]>setup/60.+1e-10 and (not ischange or day_changes[days[st]]<maxchg[m]): break
                        st+=1
                    if st>=T: continue
                    goal=rem[p]
                    # Diversify block sizes without weakening minimum-run rules.
                    if goal>need+POSITIVE_T and need>0 and np.random.random()<0.23:
                        goal=need+(goal-need)*np.random.random()
                    temp=np.zeros(T); dy=np.zeros(np.max(days)+1); total=0.; end=-1
                    for t in range(st,T):
                        avail=hours[m,t]-(setup/60. if t==st else 0.)
                        if avail<=0: continue
                        qt=min(avail*rate[m,p],daily[m,p]-dy[days[t]],goal-total)
                        if qt<POSITIVE_T: continue
                        temp[t]=qt; total+=qt; dy[days[t]]+=qt; end=t
                        if total>=goal-TOL: break
                    if total<POSITIVE_T: continue
                    if total+TOL<need:
                        if not (allow_final and abs(total-rem[p])<=TOL and st>lastend[p]): continue
                    for t in range(st,end+1):
                        if temp[t]>0: X[n,m,t]=p; Q[n,m,t]=temp[t]
                    rem[p]=max(0.,rem[p]-total)
                    lastend[p]=max(lastend[p],end)
                    if ischange: day_changes[days[st]]+=1
                    used[m,p]=1; done[p]=1; last=p; first=False; cursor=end+1
                    chosen=True
                    break
                if not chosen: break
    return X,Q


def constructive(d: Data, attempts: int, seed: int, events: list, pool_limit: int=300) -> list[Plan]:
    days=np.array([v[0] for v in d.slots],dtype=np.int64)
    args=(d.rate,d.minimum,d.daily,d.prep,d.hours,d.target,d.predecessors,d.initial,
          d.initial_qty,d.prepared,d.max_changes,days,d.family_index,d.priority,
          d.settings['AllowFinalRunBelowMinimum'],d.settings['CriticalProductFirst'],
          d.product_limits,d.family_limits)
    if NUMBA_AVAILABLE:
        start=time.perf_counter(); _construct_batch(1,seed,*args)
        events.append({'stage':'JIT compile / warm-up (not search)','runtime_sec':time.perf_counter()-start,'status':'COMPLETED','actual_attempts':0})
    start=time.perf_counter(); archive=[]; baseline=float('inf'); valid=0
    completed=0; batchsize=256
    while completed<attempts:
        n=min(batchsize,attempts-completed)
        X,Q=_construct_batch(n,(seed+completed+17)%2147483647,*args)
        # Vectorized early tonnage filter; rich checks only for contenders.
        produced=Q.sum(axis=(1,2)); unmet=d.target.sum()-produced
        for i in range(n):
            if unmet[i]>baseline+d.settings['UnmetTolerance_t']+TOL: continue
            plan=Plan(X[i].astype(np.int64),Q[i].copy(),'CONSTRUCTIVE')
            if validate(d,plan): continue
            valid+=1; met=plan.metrics; key=preference_key(d,met)
            if met['unmet']<baseline:
                baseline=met['unmet']
                archive=[p for p in archive if p.metrics['unmet']<=baseline+d.settings['UnmetTolerance_t']+TOL]
            # Pareto archive: lower unmet and lexicographically better preferences.
            if any(p.metrics['unmet']<=met['unmet']+1e-8 and preference_key(d,p.metrics)<=key for p in archive): continue
            archive=[p for p in archive if not (met['unmet']<=p.metrics['unmet']+1e-8 and key<=preference_key(d,p.metrics))]
            archive.append(plan)
            if len(archive)>pool_limit:
                # Heuristic archive cap is explicitly reported, never an optimality claim.
                archive.sort(key=lambda p:p.metrics['unmet']); keep=[archive[0]]
                keep+=sorted(archive[1:],key=lambda p:preference_key(d,p.metrics))[:pool_limit-1]
                archive=keep
        completed+=n
        if completed==attempts or completed%max(batchsize,(attempts//5//batchsize)*batchsize)==0:
            LOG.info('Heuristic %s/%s attempts; best unmet %.6f t',completed,attempts,baseline)
    events.append({'stage':'Constructive heuristic','runtime_sec':time.perf_counter()-start,
        'status':'COMPLETED','requested_attempts':attempts,'actual_attempts':completed,
        'notes':f'Independent-validated contenders={valid}; Pareto archive={len(archive)}; duplicates possible; archive cap={pool_limit}'})
    return archive


class LinearModel:
    """Small sparse MILP builder with readable constraint coefficients."""
    def __init__(self):
        self.lo=[]; self.hi=[]; self.integer=[]; self.names=[]
        self.rows=[]; self.lower=[]; self.upper=[]
    def var(self,name,lo=0.,hi=1.,integer=False):
        i=len(self.lo); self.lo.append(lo); self.hi.append(hi); self.integer.append(int(integer)); self.names.append(name); return i
    def add(self,coeff,lo=-np.inf,hi=np.inf):
        row={}
        for i,v in coeff.items() if isinstance(coeff,dict) else coeff:
            row[i]=row.get(i,0.)+v
        self.rows.append({i:v for i,v in row.items() if v}); self.lower.append(lo); self.upper.append(hi)
    def array(self,obj):
        v=np.zeros(len(self.lo))
        for i,k in obj.items(): v[i]=k
        return v
    def solve(self,obj,seconds,extra,disp=False):
        rows=self.rows+[r[0] for r in extra]
        lo=self.lower+[r[1] for r in extra]; hi=self.upper+[r[2] for r in extra]
        ii=[]; jj=[]; vv=[]
        for r,row in enumerate(rows):
            for c,v in row.items(): ii.append(r); jj.append(c); vv.append(v)
        A=coo_matrix((vv,(ii,jj)),shape=(len(rows),len(self.lo))).tocsc()
        return milp(self.array(obj),integrality=np.asarray(self.integer),
            bounds=Bounds(self.lo,self.hi),constraints=LinearConstraint(A,lo,hi),
            options={'time_limit':max(.01,seconds),'mip_rel_gap':0.,'presolve':True,'disp':disp})


@dataclass
class ModelPack:
    model: LinearModel
    q: dict
    x: dict
    y: dict
    st: dict
    objectives: dict
    unmet: dict
    first: dict
    last: dict


def make_model(d: Data) -> ModelPack:
    """Unrestricted slot MILP. Machine assignments and block order are free.

    Pairwise disjunctive run intervals enforce no-return across idle slots.
    Quantities are continuous tons; binary assignment activation is >=0.1 kg.
    """
    b=LinearModel(); q={}; x={}; y={}; st={}; first={}; last={}; exception={}; chg={}; total={}; amount={}
    T=d.T; big=T+1
    pairs=[(m,p) for m in range(d.M) for p in range(d.P) if d.rate[m,p]>0 and d.target[p]>=POSITIVE_T]
    mp_for_p={p:[(m,p) for m in range(d.M) if (m,p) in pairs] for p in range(d.P)}
    used_m={}
    for m,p in pairs:
        y[m,p]=b.var(f'y_{m}_{p}',integer=True)
        first[m,p]=b.var(f'first_{m}_{p}',hi=T-1,integer=True)
        last[m,p]=b.var(f'last_{m}_{p}',hi=T-1,integer=True)
        exception[m,p]=b.var(f'final_{m}_{p}',hi=d.settings['AllowFinalRunBelowMinimum'],integer=True)
        amount[m,p]=b.var(f'amount_{m}_{p}',hi=d.target[p])
        starts=[]; xs=[]; qs=[]; earlier=[]
        for t in range(T):
            if d.hours[m,t]<=0: continue
            k=(m,p,t)
            q[k]=b.var(f'q_{m}_{p}_{t}',hi=min(d.target[p],d.rate[m,p]*d.hours[m,t],d.daily[m,p]))
            x[k]=b.var(f'x_{m}_{p}_{t}',integer=True)
            st[k]=b.var(f'start_{m}_{p}_{t}',integer=True)
            setup=0. if d.initial[m]==p and d.prepared[m] else d.prep[m,p]
            b.add({q[k]:1,x[k]:-min(d.target[p],d.rate[m,p]*d.hours[m,t],d.daily[m,p])},hi=0)
            b.add({q[k]:1,x[k]:-POSITIVE_T},lo=0)
            b.add({q[k]:1/d.rate[m,p],st[k]:setup/60.,x[k]:-d.hours[m,t]},hi=0)
            b.add({st[k]:1,x[k]:-1},hi=0)
            b.add([(a,1) for a in earlier]+[(st[k],T)],hi=T)
            b.add({first[m,p]:1,x[k]:T},hi=t+T)
            b.add({last[m,p]:1,x[k]:-T},lo=t-T)
            starts.append(st[k]); xs.append(x[k]); qs.append(q[k]); earlier.append(x[k])
        b.add([(i,1) for i in starts]+[(y[m,p],-1)],lo=0,hi=0)
        b.add([(i,1) for i in xs]+[(y[m,p],-T)],hi=0)
        b.add([(st[m,p,t],t) for t in range(T) if (m,p,t) in st]+[(first[m,p],-1)],lo=0,hi=0)
        b.add({last[m,p]:1,y[m,p]:-(T-1)},hi=0)
        b.add([(i,1) for i in qs]+[(amount[m,p],-1)],lo=0,hi=0)
        b.add({exception[m,p]:1,y[m,p]:-1},hi=0)
        carry=d.initial_qty[m] if d.initial[m]==p else 0.
        needed=max(0.,d.minimum[m,p]-carry)
        b.add({amount[m,p]:1,y[m,p]:-needed,exception[m,p]:needed},lo=0)
        if needed<=POSITIVE_T:
            b.add({exception[m,p]:1},hi=0)
        else:
            # e=1 implies a genuinely below-minimum run, not a gratuitous exception.
            b.add({amount[m,p]:1,exception[m,p]:d.target[p]},hi=needed-POSITIVE_T+d.target[p])
        for day in sorted(set(a for a,_ in d.slots)):
            b.add({q[m,p,t]:1 for t,(dd,_) in enumerate(d.slots) if dd==day and (m,p,t) in q},hi=d.daily[m,p])
    unmet={}
    for p in range(d.P):
        unmet[p]=b.var(f'unmet_{p}',hi=0. if d.hard[p] else d.target[p])
        b.add([(amount[k],1) for k in mp_for_p[p]]+[(unmet[p],1)],lo=d.target[p],hi=d.target[p])
        b.add({exception[k]:1 for k in mp_for_p[p]},hi=1)
        for k in mp_for_p[p]:
            b.add({unmet[p]:1,exception[k]:d.target[p]},hi=d.target[p])
            for other in mp_for_p[p]:
                if other==k: continue
                # Every other run of p must finish in an earlier slot than the final remainder.
                b.add({last[other]:1,first[k]:-1,exception[k]:big,y[other]:big},hi=2*big-1)
    for m in range(d.M):
        mps=[(mm,p) for mm,p in pairs if mm==m]
        used_m[m]=b.var(f'used_machine_{m}',integer=True)
        for k in mps: b.add({y[k]:1,used_m[m]:-1},hi=0)
        b.add([(y[k],1) for k in mps]+[(used_m[m],-1)],lo=0)
        # Initial empty-machine setup does not count as a changeover.
        initial_starts=[]; earlier=[]
        for t in range(T):
            xt=[x[m,p,t] for _,p in mps if (m,p,t) in x]
            stt=[st[m,p,t] for _,p in mps if (m,p,t) in st]
            b.add({i:1 for i in xt},hi=1)
            chg[m,t]=b.var(f'change_{m}_{t}',integer=True)
            coeff=[(i,1) for i in stt]+[(chg[m,t],-1)]
            if d.initial[m]<0:
                a=b.var(f'initial_setup_{m}_{t}',integer=True)
                initial_starts.append(a)
                b.add([(a,1)]+[(i,-1) for i in stt],hi=0)
                b.add([(i,1) for i in earlier]+[(a,T)],hi=T)
                coeff.append((a,-1))
            elif (m,int(d.initial[m]),t) in st:
                coeff.append((st[m,int(d.initial[m]),t],-1))
            b.add(coeff,lo=0,hi=0)
            earlier.extend(xt)
        if d.initial[m]<0:
            b.add([(i,1) for i in initial_starts]+[(used_m[m],-1)],lo=0,hi=0)
        for day in sorted(set(a for a,_ in d.slots)):
            b.add({chg[m,t]:1 for t,(dd,_) in enumerate(d.slots) if day==dd},hi=d.max_changes[m])
        for i,k in enumerate(mps):
            for other in mps[i+1:]:
                # Exactly one ordering when both products use this machine.
                o=b.var(f'order_{m}_{k[1]}_{other[1]}',integer=True)
                b.add({last[k]:1,first[other]:-1,y[k]:big,y[other]:big,o:big},hi=3*big-1)
                b.add({last[other]:1,first[k]:-1,y[k]:big,y[other]:big,o:-big},hi=2*big-1)
        for a in range(d.P):
            for p in range(d.P):
                if d.predecessors[a,p] and (m,a) in y and (m,p) in y:
                    b.add({last[m,a]:1,first[m,p]:-1,y[m,a]:big,y[m,p]:big},hi=2*big-1)
        if d.initial[m]>=0:
            p0=int(d.initial[m])
            for _,p in mps:
                if d.predecessors[p,p0]: b.add({y[m,p]:1},hi=0)
                if p!=p0 and (m,p0) in y:
                    b.add({last[m,p0]:1,first[m,p]:-1,y[m,p0]:big,y[m,p]:big},hi=2*big-1)
    product_used={}; family_used={}; fm={}
    for p in range(d.P):
        product_used[p]=b.var(f'produced_product_{p}',integer=True)
        ys=[y[k] for k in mp_for_p[p]]
        for v in ys: b.add({v:1,product_used[p]:-1},hi=0)
        b.add([(v,1) for v in ys]+[(product_used[p],-1)],lo=0)
        b.add({v:1 for v in ys},hi=int(d.product_limits[p]))
    for f in range(len(d.families)):
        family_used[f]=b.var(f'produced_family_{f}',integer=True)
        for m in range(d.M):
            fm[m,f]=b.var(f'family_machine_{m}_{f}',integer=True)
            ys=[y[m,p] for p in range(d.P) if d.family_index[p]==f and (m,p) in y]
            for v in ys: b.add({v:1,fm[m,f]:-1},hi=0)
            b.add([(v,1) for v in ys]+[(fm[m,f],-1)],lo=0)
            b.add({fm[m,f]:1,family_used[f]:-1},hi=0)
        b.add([(fm[m,f],1) for m in range(d.M)]+[(family_used[f],-1)],lo=0)
        b.add({fm[m,f]:1 for m in range(d.M)},hi=int(d.family_limits[f]))
    product_split={v:1 for v in y.values()}; product_split.update({v:-1 for v in product_used.values()})
    family_split={v:1 for v in fm.values()}; family_split.update({v:-1 for v in family_used.values()})
    objectives={
        'MIN_UNMET':{v:1 for v in unmet.values()},
        'MIN_PRODUCT_MACHINE_SPLITS':product_split,'MIN_CHANGEOVERS':{v:1 for v in chg.values()},
        'MIN_FAMILY_MACHINE_SPLITS':family_split,'MIN_PRODUCTS_PER_MACHINE':{v:1 for v in y.values()},
        'MIN_SURPLUS':{},
        'MIN_MACHINE_PRIORITY_COST':{amount[m,p]:d.priority[m,p]-1 for m,p in pairs},
        'TIE_MASTER_PRIORITY':{amount[m,p]:d.machines[m]['Priority']-1 for m,p in pairs},
    }
    return ModelPack(b,q,x,y,st,objectives,unmet,first,last)


def decode(d: Data, pack: ModelPack, values: np.ndarray, method: str) -> Plan:
    X=np.full((d.M,d.T),-1,dtype=np.int64); Q=np.zeros((d.M,d.T))
    for k,ix in pack.x.items():
        if values[ix]>.5:
            m,p,t=k
            if X[m,t]>=0: raise RuntimeError('MILP decoder: multiple products in a shift.')
            X[m,t]=p; Q[m,t]=max(0.,float(values[pack.q[k]]))
    return Plan(X,Q,method)


def milp_hybrid(d: Data, archive: list[Plan], budget: float, events: list, solver_log: bool=False) -> tuple[Plan,Plan,dict]:
    """Incumbent-preserving hybrid, not a claimed native MIP warm start.

    scipy.optimize.milp has no public x0. Known validated plans are retained
    externally and used as objective cutoffs. Reallocation is unrestricted.
    """
    if not archive:
        empty=Plan(np.full((d.M,d.T),-1,dtype=int),np.zeros((d.M,d.T)),'EMPTY_VALID_PLAN')
        if not validate(d,empty): archive=[empty]
    t0=time.perf_counter(); pack=make_model(d)
    events.append({'stage':'MILP model build','runtime_sec':time.perf_counter()-t0,'status':'COMPLETED',
                   'notes':f'{len(pack.model.lo)} variables; {len(pack.model.rows)} rows; unrestricted assignment/order'})
    meta={'native_mip_start':False,'solver':'SciPy / HiGHS','baseline_optimal':False,
          'preference_optimal':False,'milp_solve_budget_sec':budget,'stages':[]}
    spent=0.; baseline=min(archive,key=lambda p:p.metrics['unmet']) if archive else None
    lookup={'MIN_UNMET':'unmet','MIN_PRODUCT_MACHINE_SPLITS':'product_splits','MIN_CHANGEOVERS':'changeovers',
        'MIN_FAMILY_MACHINE_SPLITS':'family_splits','MIN_PRODUCTS_PER_MACHINE':'assignments','MIN_SURPLUS':'surplus',
        'MIN_MACHINE_PRIORITY_COST':'priority_cost','TIE_MASTER_PRIORITY':'master_priority_cost'}
    def call(oid,seconds,extra):
        nonlocal spent
        if seconds<.02: return None,None
        started=time.perf_counter()
        res=pack.model.solve(pack.objectives[oid],seconds,extra,solver_log)
        elapsed=time.perf_counter()-started; spent+=elapsed
        entry={'stage':'MILP '+oid,'runtime_sec':elapsed,'status':str(res.message),
               'objective':oid,'solver_status_code':int(res.status),
               'objective_value':None if getattr(res,'fun',None) is None else float(res.fun),
               'best_bound':None if getattr(res,'mip_dual_bound',None) is None else float(res.mip_dual_bound),
               'gap':None if getattr(res,'mip_gap',None) is None else float(res.mip_gap),
               'requested_sec':seconds}
        plan=None
        if getattr(res,'x',None) is not None:
            candidate=decode(d,pack,res.x,'MILP_'+oid)
            problems=validate(d,candidate)
            if not problems: plan=candidate
            else: entry['rejected_by_validator']=problems; LOG.warning('MILP candidate rejected: %s',problems[:2])
        entry['independent_validation']='PASS' if plan is not None else 'NO_VALID_CANDIDATE'
        events.append(entry); meta['stages'].append(entry)
        LOG.info('%s: %.3f s; %s; unmet=%s',oid,elapsed,res.message,plan.metrics['unmet'] if plan else 'none')
        return plan,res
    # Allocate most time to throughput, remainder to lexicographic preferences.
    extra=[]
    if baseline is not None: extra=[(pack.objectives['MIN_UNMET'],-np.inf,baseline.metrics['unmet']+1e-7)]
    candidate,res=call('MIN_UNMET',budget*.60,extra)
    if candidate is not None:
        archive.append(candidate)
        if baseline is None or candidate.metrics['unmet']<baseline.metrics['unmet']-1e-7: baseline=candidate
    if res is not None and res.status==0 and candidate is not None:
        meta['baseline_optimal']=True
    if baseline is None:
        if res is not None and res.status==2: raise InputError('No feasible schedule: MILP proved the HARD model infeasible.')
        raise InputError('No feasible schedule found within the budget. This is not proof of infeasibility. Increase MILP time or review HARD targets.')
    baseline=baseline.clone(); baseline.method+=' | BASELINE'
    cap=baseline.metrics['unmet']+d.settings['UnmetTolerance_t']; meta['frozen_unmet_cap_t']=cap
    selected=min((p for p in archive if p.metrics['unmet']<=cap+TOL),key=lambda p:preference_key(d,p.metrics)).clone()
    extra=[(pack.objectives['MIN_UNMET'],-np.inf,cap)]
    stages=[o for o in d.objectives if o!='MIN_UNMET']+['MIN_UNMET','TIE_MASTER_PRIORITY']
    all_proved=True
    for i,oid in enumerate(stages):
        # Freeze each achieved higher-priority value before moving to the next.
        if not pack.objectives[oid]:
            events.append({'stage':'MILP '+oid,'runtime_sec':0.,'status':'CONSTANT_OBJECTIVE','notes':'Surplus is forbidden.'})
            continue
        remaining=max(0.,budget-spent)
        slots_left=sum(bool(pack.objectives[o]) for o in stages[i:])
        seconds=remaining/max(1,slots_left)
        cutoff=selected.metrics[lookup[oid]]
        candidate,res=call(oid,seconds,extra+[(pack.objectives[oid],-np.inf,cutoff+1e-7)])
        if candidate is not None and candidate.metrics['unmet']<=cap+TOL:
            # Full lexicographic comparator prevents a numerically degraded incumbent.
            if preference_key(d,candidate.metrics)<preference_key(d,selected.metrics): selected=candidate
        if res is None or res.status!=0: all_proved=False
        value=selected.metrics[lookup[oid]]
        if oid in ['MIN_PRODUCT_MACHINE_SPLITS','MIN_CHANGEOVERS','MIN_FAMILY_MACHINE_SPLITS','MIN_PRODUCTS_PER_MACHINE']:
            value=round(value)
        extra.append((pack.objectives[oid],-np.inf,value+1e-7))
    meta['preference_optimal']=all_proved
    meta['milp_solve_actual_sec']=spent
    meta['status']='OPTIMAL_ALL_STAGES' if all_proved and meta['baseline_optimal'] else 'FEASIBLE_NOT_PROVEN_OPTIMAL'
    if validate(d,selected,cap): raise RuntimeError('Selected plan failed final independent validation.')
    if validate(d,baseline): raise RuntimeError('Baseline failed final independent validation.')
    return baseline,selected,meta


def detail_rows(d: Data, plan: Plan, plan_id: str) -> list[dict]:
    rows=[]; counter=0
    for m in range(d.M):
        last=int(d.initial[m]); lastday=None; first=True; run=''
        for t,(day,shift) in enumerate(d.slots):
            p=int(plan.product[m,t]); q=float(plan.quantity[m,t]); setup=0.; change=0; cross=0
            previous=d.products[last]['Product_ID'] if last>=0 else ''
            if p>=0:
                if first or p!=last:
                    counter+=1; run=f'{plan_id[-1]}{counter:03d}'
                    setup=d.prep[m,p]
                    if first and p==d.initial[m] and d.prepared[m]: setup=0.
                change=int(last>=0 and p!=last)
                cross=int(change and (lastday is None or lastday!=day))
                last=p; lastday=day; first=False
                rate=d.rate[m,p]; prod=q/rate
            else:
                rate=0.; prod=0.
            rows.append({'Plan_ID':plan_id,'Day':f'Day-{day}','Shift':shift,
                'Machine_ID':d.machines[m]['Machine_ID'],'Machine':d.machines[m]['Machine'],
                'Product_ID':d.products[p]['Product_ID'] if p>=0 else '',
                'Product':d.products[p]['Product'] if p>=0 else '',
                'Family':d.products[p]['Family'] if p>=0 else '',
                'Run_ID':run if p>=0 else '', 'Planned_t':q,'Preparation_min':setup,
                'Previous_Product_ID':previous,'Is_Changeover':change,'Cross_Day':cross,
                'Shift_Status':'PRODUCTION' if p>=0 else ('CLOSED' if not d.hours[m,t] else 'IDLE'),
                'Available_h':d.hours[m,t],'Capacity_tph':rate,'Production_h':prod,
                'Occupied_h':prod+setup/60.,'Idle_h':max(0.,d.hours[m,t]-prod-setup/60.),
                'MaxDaily_t':d.daily[m,p] if p>=0 else 0.,'m':m,'p':p,'t':t})
    dq=defaultdict(float); dc=defaultdict(int)
    for r in rows:
        dq[r['m'],r['p'],r['Day']]+=r['Planned_t']; dc[r['m'],r['Day']]+=r['Is_Changeover']
    for r in rows:
        r['Daily_Planned_t']=dq[r['m'],r['p'],r['Day']] if r['p']>=0 else 0.
        r['Daily_Changeovers']=dc[r['m'],r['Day']]
        r['Slot_Check']='PASS'
    return rows



class ExcelWriter:
    """Write formulas and verified value caches without requiring desktop Excel."""
    def __init__(self):
        self.wb=Workbook(); self.wb.remove(self.wb.active); self.cache={}
        self.wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True,forceFullCalc=True,calcMode='auto')
    def value(self,ws,row,col,value):
        cell=ws.cell(row,col,value)
        if isinstance(value,str) and value.startswith(('=','+','-','@')):
            cell.data_type='s'  # Product names cannot inject formulas.
        return cell
    def formula(self,ws,row,col,formula,value):
        c=ws.cell(row,col,formula); self.cache[(ws.title,c.coordinate)]=value; return c
    def save(self,path):
        path=Path(path); tmp=path.with_suffix('.building.xlsx')
        self.wb.save(tmp)
        ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(tmp) as src, zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as dst:
            byxml={f'xl/worksheets/sheet{i+1}.xml':ws.title for i,ws in enumerate(self.wb)}
            for info in src.infolist():
                body=src.read(info.filename)
                if info.filename in byxml:
                    root=ET.fromstring(body); sheet=byxml[info.filename]
                    for c in root.findall('.//s:c',ns):
                        key=(sheet,c.attrib['r'])
                        if key not in self.cache: continue
                        value=self.cache[key]
                        v=c.find('s:v',ns)
                        if v is None: v=ET.SubElement(c,'{'+ns['s']+'}v')
                        if value is None: value=''
                        if isinstance(value,(str,)):
                            c.set('t','str'); v.text=value
                        elif isinstance(value,(bool,np.bool_)):
                            c.set('t','b'); v.text='1' if value else '0'
                        else:
                            c.attrib.pop('t',None)
                            if not math.isfinite(float(value)): raise RuntimeError('Invalid formula cache value.')
                            v.text=format(float(value),'.16g')
                    body=ET.tostring(root,encoding='utf-8',xml_declaration=True)
                dst.writestr(info,body)
        tmp.unlink(missing_ok=True)


NAVY='17365D'; TEAL='DDEEF0'; LIGHT='F1F5FA'; ORANGE='FFF2CC'
NUMFMT='#,##0.000;[Red]-#,##0.000;0.000'


def initialize_sheet(ws,title,subtitle,ncols):
    ws.sheet_view.showGridLines=False
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
    ws.cell(1,1,title).font=Font(name='Calibri',size=16,bold=True,color='FFFFFF')
    ws.cell(1,1).fill=PatternFill('solid',fgColor=NAVY)
    ws.row_dimensions[1].height=30
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncols)
    ws.cell(2,1,subtitle).font=Font(name='Calibri',size=10,color='595959')
    ws.cell(2,1).alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[2].height=32
    for c in range(1,ncols+1): ws.column_dimensions[get_column_letter(c)].width=17
    ws.freeze_panes='C6'
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0


def header(ws,row,labels):
    for col,value in enumerate(labels,1):
        c=ws.cell(row,col,value); c.font=Font(name='Calibri',bold=True,color='FFFFFF',size=10)
        c.fill=PatternFill('solid',fgColor=NAVY)
        c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[row].height=32


def finish_table(ws,start,end,ncols):
    for r in range(start,end+1):
        ws.row_dimensions[r].height=max(ws.row_dimensions[r].height or 0,24)
        for c in range(1,ncols+1):
            z=ws.cell(r,c)
            z.alignment=Alignment(vertical='center',wrap_text=True)
            z.font=Font(name='Calibri',size=10,color='000000' if z.data_type=='f' else '00643C')
            if r%2==0: z.fill=PatternFill('solid',fgColor=LIGHT)
            if isinstance(z.value,(int,float)) or z.data_type=='f': z.number_format=NUMFMT


def export_excel(d: Data, baseline: Plan, selected: Plan, meta: dict, events: list, config: dict, path: Path) -> None:
    """Five visible sheets. All records and input snapshots live in Detaylar."""
    ew=ExcelWriter(); wb=ew.wb
    dash=wb.create_sheet('dashboard'); rp=wb.create_sheet('results-plan')
    ps=wb.create_sheet('product-summary'); ms=wb.create_sheet('machine-summary'); dt=wb.create_sheet('Detaylar')
    initialize_sheet(dt,'DETAYLAR | V6 HYBRID','Selected and baseline details, validation, objectives and input snapshot.',24)
    dt.freeze_panes='F7'
    records={}; ranges={}; row=6
    names=['Plan_ID','Day','Shift','Machine_ID','Machine','Product_ID','Product','Family','Run_ID','Planned_t',
        'Preparation_min','Previous_Product_ID','Is_Changeover','Cross_Day','Shift_Status','Available_h',
        'Capacity_tph','Production_h','Occupied_h','Idle_h','MaxDaily_t','Daily_Planned_t','Daily_Changeovers','Slot_Check']
    section_links=[]
    def section(title,labels):
        nonlocal row
        row+=2; title_row=row
        dt.merge_cells(start_row=row,start_column=1,end_row=row,end_column=max(4,len(labels)))
        dt.cell(row,1,title).font=Font(name='Calibri',bold=True,color=NAVY,size=12)
        dt.cell(row,1).fill=PatternFill('solid',fgColor=TEAL); dt.row_dimensions[row].height=28
        row+=1; header(dt,row,labels); row+=1
        section_links.append((title,title_row))
        return row
    for pid,plan in [('PLAN_SELECTED',selected),('PLAN_BASELINE',baseline)]:
        rr=detail_rows(d,plan,pid); records[pid]=rr
        start=section(pid+' | SHIFT RECORDS',names)
        for rec in rr:
            r=row
            for j,n in enumerate(names,1): ew.value(dt,r,j,rec[n])
            ew.formula(dt,r,18,f'=IF(Q{r}>0,J{r}/Q{r},0)',rec['Production_h'])
            ew.formula(dt,r,19,f'=R{r}+K{r}/60',rec['Occupied_h'])
            ew.formula(dt,r,20,f'=MAX(P{r}-S{r},0)',rec['Idle_h'])
            rec['_excel_row']=r; row+=1
        end=row-1; ranges[pid]=(start,end)
        for rec in rr:
            r=rec['_excel_row']
            if rec['p']>=0:
                ew.formula(dt,r,22,f'=SUMIFS(J${start}:J${end},D${start}:D${end},D{r},F${start}:F${end},F{r},B${start}:B${end},B{r})',rec['Daily_Planned_t'])
            ew.formula(dt,r,23,f'=SUMIFS(M${start}:M${end},D${start}:D${end},D{r},B${start}:B${end},B{r})',rec['Daily_Changeovers'])
        finish_table(dt,start,end,len(names)); dt.row_dimensions.group(start,end,outline_level=1,hidden=False)
    selected_records=records['PLAN_SELECTED']; a,z=ranges['PLAN_SELECTED']
    # Machine-product assignments provide formula-based distinct counts.
    ar={}; famr={}; assignment={}
    ast=section('SELECTED | MACHINE-PRODUCT ASSIGNMENTS',['Machine_ID','Machine','Product_ID','Product','Family','Planned_t','Assigned','Rate_tph','MinRun_t','MaxDaily_t','Preparation_min','MachinePriority','Priority_cost'])
    for m in range(d.M):
        for p in range(d.P):
            r=row; ar[m,p]=r
            qv=sum(rec['Planned_t'] for rec in selected_records if rec['m']==m and rec['p']==p)
            assignment[m,p]=int(qv>TOL)
            vals=[d.machines[m]['Machine_ID'],d.machines[m]['Machine'],d.products[p]['Product_ID'],d.products[p]['Product'],d.products[p]['Family'],qv,int(qv>TOL),d.rate[m,p],d.minimum[m,p],d.daily[m,p],d.prep[m,p],d.priority[m,p],qv*max(0,d.priority[m,p]-1)]
            for j,v in enumerate(vals,1): ew.value(dt,r,j,v)
            ew.formula(dt,r,6,f'=SUMIFS(J${a}:J${z},D${a}:D${z},A{r},F${a}:F${z},C{r})',qv)
            ew.formula(dt,r,7,f'=IF(F{r}>{TOL},1,0)',int(qv>TOL))
            ew.formula(dt,r,13,f'=F{r}*MAX(L{r}-1,0)',qv*max(0,d.priority[m,p]-1))
            row+=1
    finish_table(dt,ast,row-1,13); dt.row_dimensions.group(ast,row-1,outline_level=1,hidden=True)
    family_start=section('SELECTED | FAMILY SUMMARY',['Family','Target_Products','Produced_Products','Target_t','Planned_t','Unmet_t','Machine_Count','Machines_Used','Extra_Machines','Target_Status'])
    for f,fname in enumerate(d.families):
        pp=[p for p in range(d.P) if d.family_index[p]==f]
        mm=[m for m in range(d.M) if any(assignment[m,p] for p in pp)]
        qv=sum(r['Planned_t'] for r in selected_records if r['Family']==fname)
        tv=float(d.target[pp].sum()); count=sum(any(assignment[m,p] for m in range(d.M)) for p in pp)
        vals=[fname,len(pp),count,tv,qv,max(0,tv-qv),len(mm),' '.join(d.machines[m]['Machine'] for m in mm),max(len(mm)-1,0),'MET' if tv-qv<=TOL else 'SOFT_UNMET']
        for j,v in enumerate(vals,1): ew.value(dt,row,j,v)
        ew.formula(dt,row,5,f'=SUMIF(H${a}:H${z},A{row},J${a}:J${z})',qv)
        ew.formula(dt,row,6,f'=MAX(D{row}-E{row},0)',max(0,tv-qv))
        ew.formula(dt,row,9,f'=MAX(G{row}-1,0)',max(len(mm)-1,0)); famr[f]=row; row+=1
    finish_table(dt,family_start,row-1,10)
    runstart=section('SELECTED | RUN SUMMARY',['Run_ID','Machine','Product_ID','Product','Start_Day','Start_Shift','End_Day','End_Shift','Planned_t','MinRun_t','Carried_Run_t','Preparation_min','Status'])
    runlist=list(dict.fromkeys(r['Run_ID'] for r in selected_records if r['Run_ID']))
    for run in runlist:
        rr=[r for r in selected_records if r['Run_ID']==run]; f,l=rr[0],rr[-1]; m,p=f['m'],f['p']
        qv=sum(r['Planned_t'] for r in rr); prepv=sum(r['Preparation_min'] for r in rr); carry=d.initial_qty[m] if d.initial[m]==p else 0.
        status='OK' if qv+carry+TOL>=d.minimum[m,p] else 'FINAL_RUN_ALLOWED'
        vals=[run,f['Machine'],f['Product_ID'],f['Product'],f['Day'],f['Shift'],l['Day'],l['Shift'],qv,d.minimum[m,p],carry,prepv,status]
        for j,v in enumerate(vals,1): ew.value(dt,row,j,v)
        ew.formula(dt,row,9,f'=SUMIF(I${a}:I${z},A{row},J${a}:J${z})',qv)
        ew.formula(dt,row,12,f'=SUMIF(I${a}:I${z},A{row},K${a}:K${z})',prepv); row+=1
    finish_table(dt,runstart,row-1,13)
    cs=section('SELECTED | CHANGEOVER EVENTS',['Machine','Day','Shift','Previous_Product_ID','New_Product_ID','Product','Cross_Day','Preparation_min','Daily_Count','Daily_Limit','Status'])
    for rec in selected_records:
        if rec['Is_Changeover']:
            vals=[rec['Machine'],rec['Day'],rec['Shift'],rec['Previous_Product_ID'],rec['Product_ID'],rec['Product'],rec['Cross_Day'],rec['Preparation_min'],rec['Daily_Changeovers'],int(d.max_changes[rec['m']]),'PASS']
            for j,v in enumerate(vals,1): ew.value(dt,row,j,v)
            row+=1
    finish_table(dt,cs,row-1,11)
    vs=section('INDEPENDENT VALIDATION',['Plan_ID','Check','Status','Violations','Evidence'])
    checks={'ELIGIBILITY':'Machine-product compatibility','SHIFT_TIME':'Shift time including initial/setup',
        'DAILY_TON':'Machine-product DAILY limit','CALENDAR':'Calendar and one product per shift',
        'DAILY_CHANGEOVER':'Daily changes including cross-day','SEQUENCE':'Explicit SAME_MACHINE pairs',
        'NO_RETURN':'No A-B-A; idle does not reset','MIN_RUN':'MinRun or chronological final remainder',
        'HARD_TARGET':'HARD targets','SURPLUS':'No surplus','ALLOCATION_PRODUCT':'Product hard limits',
        'ALLOCATION_FAMILY':'Family hard limits','UNMET_CAP':'Frozen total unmet allowance'}
    for pid,pl in [('PLAN_BASELINE',baseline),('PLAN_SELECTED',selected)]:
        bad=validate(d,pl,meta['frozen_unmet_cap_t'] if pid=='PLAN_SELECTED' else None)
        if bad: raise RuntimeError(f'Refusing to export invalid plan: {bad}')
        for code,label in checks.items():
            for j,v in enumerate([pid,label,'PASS',0,'Recomputed from raw slot records'],1): ew.value(dt,row,j,v)
            row+=1
    finish_table(dt,vs,row-1,5)
    excstart=section('SELECTED | UNMET TARGETS / NOTES',['Product_ID','Product','TargetMode','Target_t','Planned_t','Unmet_t','Reason'])
    for p in range(d.P):
        qv=sum(r['Planned_t'] for r in selected_records if r['p']==p)
        if d.target[p]-qv>TOL:
            vals=[d.products[p]['Product_ID'],d.products[p]['Product'],'HARD' if d.hard[p] else 'SOFT',d.target[p],qv,max(0,d.target[p]-qv),'Unmet in selected schedule; no causal bottleneck proof claimed.']
            for j,v in enumerate(vals,1): ew.value(dt,row,j,v)
            row+=1
    finish_table(dt,excstart,row-1,7)
    compstart=section('COMPARISON',['Plan_ID','Method','Target_t','Planned_t','Unmet_t','Surplus_t','Split_Products','Extra_Product_Machines','Changeovers','Extra_Family_Machines','Runs','Priority_Cost','Validity'])
    compr={}
    for pid,pl in [('PLAN_BASELINE',baseline),('PLAN_SELECTED',selected)]:
        m=pl.metrics; st,en=ranges[pid]; compr[pid]=row
        vals=[pid,pl.method,m['target'],m['planned'],m['unmet'],m['surplus'],m['split_products'],m['product_splits'],m['changeovers'],m['family_splits'],m['runs'],m['priority_cost'],'FEASIBLE']
        for j,v in enumerate(vals,1): ew.value(dt,row,j,v)
        ew.formula(dt,row,4,f'=SUM(J{st}:J{en})',m['planned'])
        ew.formula(dt,row,5,f'=MAX(C{row}-D{row},0)',m['unmet'])
        ew.formula(dt,row,9,f'=SUM(M{st}:M{en})',m['changeovers']); row+=1
    finish_table(dt,compstart,row-1,13)
    rs=section('RUNTIME | measured stages; not estimates',['Stage','Runtime_sec','Status','Requested_Attempts','Actual_Attempts','Requested_sec','Objective','Best_Bound','Gap','Notes'])
    for ev in events:
        vals=[ev.get('stage',''),ev.get('runtime_sec',0),ev.get('status',''),ev.get('requested_attempts',''),ev.get('actual_attempts',''),ev.get('requested_sec',''),ev.get('objective',''),ev.get('best_bound',''),ev.get('gap',''),ev.get('notes','')]
        for j,v in enumerate(vals,1): ew.value(dt,row,j,v if v is not None else '')
        row+=1
    finish_table(dt,rs,row-1,10)
    metadata_start=section('EXECUTION METADATA',['Parameter','Value'])
    values={'CodeVersion':VERSION,'SourceInput':d.path.name,'Input_SHA256':d.sha256,
        'Python':platform.python_version(),'NumPy':np.__version__,'SciPy':scipy.__version__,
        'NumbaAvailable':str(NUMBA_AVAILABLE),'Platform':platform.platform(),
        'Status':meta['status'],'BaselineOptimalProved':str(meta['baseline_optimal']),
        'PreferenceStagesProved':str(meta['preference_optimal']),
        'IncumbentPolicy':'Externally retained; objective cutoff. scipy.milp does not expose x0.',
        'UnmetTolerance_t':d.settings['UnmetTolerance_t'],'FrozenUnmetCap_t':meta['frozen_unmet_cap_t'],
        'NumericActivation_t':POSITIVE_T,'ValidationTolerance':TOL,
        'FinalRunTiming':'Final remainder starts after every other same-product run ends in earlier slots.',
        **{f'CLI_{k}':str(v) for k,v in config.items()}}
    for k,v in values.items(): ew.value(dt,row,1,k); ew.value(dt,row,2,v); row+=1
    finish_table(dt,metadata_start,row-1,2)
    for sh,rr in d.tables.items():
        start=section('INPUT SNAPSHOT | '+sh,SCHEMA[sh])
        for rec in rr:
            for j,k in enumerate(SCHEMA[sh],1): ew.value(dt,row,j,rec[k])
            row+=1
        finish_table(dt,start,row-1,len(SCHEMA[sh])); dt.row_dimensions.group(start,row-1,outline_level=1,hidden=True)
    # Keep detail text readable without widening numeric columns everywhere.
    for col,width in {'A':25,'B':24,'D':23,'E':20,'F':20,'G':32,'H':25,'I':20,'J':23,'X':18}.items(): dt.column_dimensions[col].width=width
    dt.row_dimensions[4].height=28
    for j,(name,r) in enumerate(section_links[:8],1):
        c=dt.cell(4,j,name.split('|')[0].strip()); c.hyperlink=f"#'Detaylar'!A{r}"; c.style='Hyperlink'; c.alignment=Alignment(wrap_text=True)
    # Product summary: formulas reference the same selected slot records.
    initialize_sheet(ps,'PRODUCT SUMMARY | SELECTED PLAN','Target status and machine allocation are separate. Quantities are tons.',13)
    ph=['Product_ID','Product','Family','TargetMode','Target_t','Planned_t','Unmet_t','Surplus_t','Fulfillment','Machines','Machine_Count','Runs','Target_Status']
    header(ps,5,ph)
    for p,prod in enumerate(d.products):
        r=6+p; qv=sum(rec['Planned_t'] for rec in selected_records if rec['p']==p); mm=[m for m in range(d.M) if assignment[m,p]]
        unmet=max(0.,d.target[p]-qv); status='MET' if unmet<=TOL else ('HARD_UNMET' if d.hard[p] else 'SOFT_UNMET')
        vals=[prod['Product_ID'],prod['Product'],prod['Family'],'HARD' if d.hard[p] else 'SOFT',d.target[p],qv,unmet,max(qv-d.target[p],0),min(qv/d.target[p],1) if d.target[p]>0 else 1,' '.join(d.machines[m]['Machine'] for m in mm),len(mm),len(mm),status]
        for j,v in enumerate(vals,1): ew.value(ps,r,j,v)
        ew.formula(ps,r,6,f'=SUMIF(Detaylar!F${a}:F${z},A{r},Detaylar!J${a}:J${z})',qv)
        ew.formula(ps,r,7,f'=MAX(E{r}-F{r},0)',unmet)
        ew.formula(ps,r,8,f'=MAX(F{r}-E{r},0)',max(qv-d.target[p],0))
        ew.formula(ps,r,9,f'=IF(E{r}=0,1,MIN(F{r}/E{r},1))',vals[8])
        countformula='+' .join(f'Detaylar!G{ar[m,p]}' for m in range(d.M))
        ew.formula(ps,r,11,'='+countformula,len(mm)); ew.formula(ps,r,12,f'=K{r}',len(mm))
        ew.formula(ps,r,13,f'=IF(G{r}<={TOL},"MET",IF(D{r}="HARD","HARD_UNMET","SOFT_UNMET"))',status)
        chunks=[f'IF(Detaylar!G{ar[m,p]}=1,"{d.machines[m]["Machine"].replace(chr(34),chr(34)*2)} ","")' for m in range(d.M)]
        ew.formula(ps,r,10,'=TRIM('+'&'.join(chunks)+')',vals[9])
    finish_table(ps,6,5+d.P,len(ph)); totalrow=6+d.P
    ps.cell(totalrow,2,'TOTAL').font=Font(bold=True)
    for c,k in [(5,'target'),(6,'planned'),(7,'unmet'),(8,'surplus')]:
        ew.formula(ps,totalrow,c,f'=SUM({get_column_letter(c)}6:{get_column_letter(c)}{totalrow-1})',selected.metrics[k])
        ps.cell(totalrow,c).number_format=NUMFMT
    ps.column_dimensions['B'].width=32; ps.column_dimensions['C'].width=20; ps.column_dimensions['J'].width=24
    ps.auto_filter.ref=f'A5:M{totalrow-1}'
    for r in range(6,totalrow):
        ps.cell(r,9).number_format='0.0%'
        ps.cell(r,11).number_format='#,##0'
        ps.cell(r,12).number_format='#,##0'
    # Machine summary.
    initialize_sheet(ms,'MACHINE SUMMARY | SELECTED PLAN','Production hours + actual preparation hours = occupied hours. Initial setup consumes time.',13)
    mh=['Machine_ID','Machine','Planned_t','Available_h','Production_h','Preparation_h','Occupied_h','Idle_h','Utilization','Products','Runs','Changeovers','Max_Daily_Changes']
    header(ms,5,mh)
    for m in range(d.M):
        r=6+m; rr=[v for v in selected_records if v['m']==m]
        vals=[d.machines[m]['Machine_ID'],d.machines[m]['Machine'],sum(v['Planned_t'] for v in rr),sum(v['Available_h'] for v in rr),sum(v['Production_h'] for v in rr),sum(v['Preparation_min'] for v in rr)/60.,sum(v['Occupied_h'] for v in rr),sum(v['Idle_h'] for v in rr),0.,sum(assignment[m,p] for p in range(d.P)),sum(assignment[m,p] for p in range(d.P)),sum(v['Is_Changeover'] for v in rr),max(v['Daily_Changeovers'] for v in rr)]
        vals[8]=vals[6]/vals[3] if vals[3]>0 else 0.
        for j,v in enumerate(vals,1): ew.value(ms,r,j,v)
        for c,source in [(3,'J'),(4,'P'),(5,'R'),(6,'K'),(7,'S'),(8,'T'),(12,'M')]:
            formula=f'=SUMIF(Detaylar!D${a}:D${z},A{r},Detaylar!{source}${a}:{source}${z})'+('/60' if c==6 else '')
            ew.formula(ms,r,c,formula,vals[c-1])
        ew.formula(ms,r,9,f'=IF(D{r}=0,0,G{r}/D{r})',vals[8])
        ew.formula(ms,r,10,'='+'+'.join(f'Detaylar!G{ar[m,p]}' for p in range(d.P)),vals[9])
        ew.formula(ms,r,11,f'=J{r}',vals[10])
    finish_table(ms,6,5+d.M,len(mh))
    for r in range(6,6+d.M):
        ms.cell(r,9).number_format='0.0%'
        for c in [10,11,12,13]: ms.cell(r,c).number_format='#,##0'
    ms.column_dimensions['B'].width=18
    # Operational matrix: original 4-column blocks per machine.
    initialize_sheet(rp,'RESULTS PLAN | SELECTED V6 PLAN','One product per machine-shift. Prep is minutes; quantities are tons. CLOSED = calendar unavailable.',2+4*d.M)
    rp.cell(4,1,'Day'); rp.cell(4,2,'Shift')
    for m in range(d.M):
        c=3+4*m; rp.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+3)
        rp.cell(4,c,d.machines[m]['Machine']).font=Font(bold=True,color='FFFFFF'); rp.cell(4,c).fill=PatternFill('solid',fgColor=NAVY)
        for j,title in enumerate(['Prep_min','Product','Planned_t','Run_ID']):
            rp.cell(5,c+j,title).font=Font(bold=True,color=NAVY)
        for j,w in enumerate([10,29,13,11]): rp.column_dimensions[get_column_letter(c+j)].width=w
    for t,(day,shift) in enumerate(d.slots):
        r=6+t; ew.value(rp,r,1,f'Day-{day}'); ew.value(rp,r,2,shift); rp.row_dimensions[r].height=38
        for m in range(d.M):
            rec=selected_records[m*d.T+t]; sr=rec['_excel_row']; c=3+4*m
            if rec['p']>=0:
                for j,source,cache in [(0,'K',rec['Preparation_min']),(1,'G',rec['Product']),(2,'J',rec['Planned_t']),(3,'I',rec['Run_ID'])]:
                    ew.formula(rp,r,c+j,f'=Detaylar!{source}{sr}',cache)
            else: ew.value(rp,r,c+1,'CLOSED' if rec['Shift_Status']=='CLOSED' else 'IDLE')
            for j in range(4):
                cell=rp.cell(r,c+j); cell.alignment=Alignment(wrap_text=True,vertical='center'); cell.font=Font(name='Calibri',size=10)
                if rec['Shift_Status']=='CLOSED': cell.fill=PatternFill('solid',fgColor='E3E3E3')
                elif rec['Preparation_min']>0: cell.fill=PatternFill('solid',fgColor=TEAL)
                elif day%2==0: cell.fill=PatternFill('solid',fgColor=LIGHT)
            rp.cell(r,c+2).number_format=NUMFMT
    rp.column_dimensions['A'].width=12; rp.column_dimensions['B'].width=8; rp.freeze_panes='C6'
    rp.print_area=f'A1:{get_column_letter(2+4*d.M)}{5+d.T}'
    # Dashboard: only links and proven status; no blanket optimality assertion.
    initialize_sheet(dash,'PRODUCTION PLANNING | V6 HYBRID','Local Python calculation. All detailed sections are on Detaylar.',6)
    header(dash,5,['Metric','Value','Note','','Objective order',''])
    rows=[('Execution_Status','COMPLETED',''),('Validation_Status','FEASIBLE','Independent slot/run checks passed'),
        ('Target_Status','ALL_TARGETS_MET' if selected.metrics['unmet']<=TOL else 'SOFT_TARGETS_UNMET','SOFT shortfalls are reported, not hidden'),
        ('Solver_Status',meta['status'],'Optimality is stage-specific; time limit is not a proof'),
        ('Total_Target_t',selected.metrics['target'],''),('Planned_t',selected.metrics['planned'],''),
        ('Unmet_t',selected.metrics['unmet'],''),('Surplus_t',selected.metrics['surplus'],''),
        ('Fulfillment',1-selected.metrics['unmet']/selected.metrics['target'] if selected.metrics['target'] else 1.,''),
        ('Runs',selected.metrics['runs'],''),('Changeovers',selected.metrics['changeovers'],'Includes previous-day transitions'),
        ('Split_Products',selected.metrics['split_products'],'Number of products using >1 machine'),
        ('Extra_Product_Machines',selected.metrics['product_splits'],'v6 objective: sum(max(machine count - 1, 0))'),
        ('Extra_Family_Machines',selected.metrics['family_splits'],''),
        ('Baseline_Unmet_t',baseline.metrics['unmet'],'Best found with same operational inputs'),
        ('Selected_Additional_Unmet_t',selected.metrics['unmet']-baseline.metrics['unmet'],'May be negative if later search also improves throughput'),
        ('UnmetTolerance_t',d.settings['UnmetTolerance_t'],'Applied once to the total plan'),
        ('Tolerance_Status','PASS','Selected unmet <= frozen baseline + tolerance'),
        ('Heuristic_Attempts',config['heuristic_trials'],''),('MILP_Budget_sec',config['milp_seconds'],''),
        ('Runtime_before_export_sec',sum(e.get('runtime_sec',0) for e in events),'Full end-to-end runtime is in the companion JSON / log'),
        ('Selected_Method',selected.method,''),('Input',d.path.name,''),('Seed',config['seed'],'Same seed fixes heuristic sampling; time-limited MILP can differ')]
    for i,(label,value,note) in enumerate(rows,6):
        ew.value(dash,i,1,label); ew.value(dash,i,2,value); ew.value(dash,i,3,note)
    for i,oid in enumerate(d.objectives,6): ew.value(dash,i,5,oid)
    for r,source in [(10,'E'),(11,'F'),(12,'G'),(13,'H')]:
        ew.formula(dash,r,2,f"='product-summary'!{source}{totalrow}",rows[r-6][1])
    ew.formula(dash,14,2,'=IF(B10=0,1,1-B12/B10)',rows[8][1])
    ew.formula(dash,15,2,f"=SUM('product-summary'!L6:L{totalrow-1})",selected.metrics['runs'])
    ew.formula(dash,16,2,f'=SUM(Detaylar!M{a}:M{z})',selected.metrics['changeovers'])
    ew.formula(dash,17,2,f'=COUNTIF(\'product-summary\'!K6:K{totalrow-1},">1")',selected.metrics['split_products'])
    ew.formula(dash,18,2,'='+'+'.join(f'MAX(\'product-summary\'!K{6+p}-1,0)' for p in range(d.P)),selected.metrics['product_splits'])
    ew.formula(dash,19,2,'='+'+'.join(f'Detaylar!I{famr[f]}' for f in range(len(d.families))),selected.metrics['family_splits'])
    ew.formula(dash,20,2,f'=Detaylar!E{compr["PLAN_BASELINE"]}',baseline.metrics['unmet'])
    ew.formula(dash,21,2,'=B12-B20',selected.metrics['unmet']-baseline.metrics['unmet'])
    ew.formula(dash,23,2,f'=IF(B21<=B22+{TOL},"PASS","FAIL")','PASS')
    finish_table(dash,6,5+len(rows),3)
    dash.cell(14,2).number_format='0.0%'
    for r in [15,16,17,18,19,24,29]: dash.cell(r,2).number_format='#,##0'
    for i in range(6,6+len(rows)):
        dash.row_dimensions[i].height=30
    dash.column_dimensions['A'].width=34; dash.column_dimensions['B'].width=40; dash.column_dimensions['C'].width=52
    dash.column_dimensions['D'].width=3; dash.column_dimensions['E'].width=39; dash.column_dimensions['F'].width=4
    for r in [10,11,12,15,16]: dash.cell(r,2).fill=PatternFill('solid',fgColor=TEAL); dash.cell(r,2).font=Font(bold=True,size=12)
    for ws in wb:
        ws.sheet_properties.outlinePr.summaryRight=False
        ws.sheet_properties.outlinePr.summaryBelow=False
        ws.print_options.horizontalCentered=True
    wb.active=1
    ew.save(path)
    # Verify formula caches exist and no invalid references were introduced.
    formulas=load_workbook(path,data_only=False); cached=load_workbook(path,data_only=True)
    for ws in formulas:
        for rr in ws:
            for cell in rr:
                if cell.data_type=='f':
                    if '#REF!' in cell.value: raise RuntimeError('Excel formula contains #REF!')
                    cv=cached[ws.title][cell.coordinate]
                    if cv.data_type=='e': raise RuntimeError(f'Excel formula error: {ws.title}!{cell.coordinate}')
    formulas.close(); cached.close()


def parse_args(argv=None):
    ap=argparse.ArgumentParser(
        description='Excel input -> hybrid production planning -> Excel output. All planning parameters are read from the input workbook.'
    )
    ap.add_argument('input', type=Path, help='v6 Excel-only input workbook (.xlsx)')
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args=parse_args(argv); start=time.perf_counter(); events=[]
    try:
        input_path=args.input.expanduser().resolve()
        if input_path.suffix.lower()!='.xlsx':
            raise InputError('Input extension must be .xlsx')
        output=input_path.with_name(input_path.stem + '_output.xlsx')
        if output==input_path:
            raise InputError('Output must not overwrite input.')
        output.parent.mkdir(parents=True,exist_ok=True)
        if output.exists():
            output.unlink()
        logging.basicConfig(level=logging.INFO,format='%(asctime)s | %(message)s',handlers=[logging.StreamHandler(sys.stdout)],force=True)
        d=load_input(input_path)
        events.append({'stage':'Excel input read / strict validation','runtime_sec':time.perf_counter()-start,'status':'PASS'})
        LOG.info('Input PASS: %s products, %s machines, %s slots per machine; target %.3f t',d.P,d.M,d.T,d.target.sum())
        for note in d.warnings: LOG.warning(note)

        trials=d.settings['HeuristicTrials']
        budget=d.settings['MILPPolishTimeSec']
        seed=d.settings['RandomSeed']
        if trials<1 or budget<0 or not math.isfinite(budget):
            raise InputError('HeuristicTrials must be positive and MILPPolishTimeSec finite/nonnegative.')

        config={
            'heuristic_trials':trials,
            'milp_seconds':budget,
            'seed':seed,
            'parameter_source':'Excel settings sheet only',
            'input_settings_overridden':False,
        }
        archive=constructive(d,trials,seed,events)
        baseline,selected,meta=milp_hybrid(d,archive,budget,events,False)

        s=time.perf_counter()
        for p in [baseline,selected]:
            bad=validate(d,p,meta['frozen_unmet_cap_t'] if p is selected else None)
            if bad:
                raise RuntimeError(f'Final plan invalid: {bad}')
        events.append({'stage':'Final independent validation','runtime_sec':time.perf_counter()-s,'status':'PASS'})

        s=time.perf_counter()
        export_excel(d,baseline,selected,meta,events,config,output)
        events.append({'stage':'Excel export / formula cache verification','runtime_sec':time.perf_counter()-s,'status':'COMPLETED'})
        meta['wall_clock_to_excel_complete_sec']=time.perf_counter()-start

        LOG.info('BASELINE | planned=%.6f unmet=%.6f',baseline.metrics['planned'],baseline.metrics['unmet'])
        LOG.info('SELECTED | planned=%.6f unmet=%.6f splits=%s changes=%s | VALIDATION PASS',
            selected.metrics['planned'],selected.metrics['unmet'],selected.metrics['product_splits'],selected.metrics['changeovers'])
        LOG.info('STATUS %s | runtime-to-Excel %.3f s | %s',meta['status'],meta['wall_clock_to_excel_complete_sec'],output)
        return 0
    except (InputError,PermissionError,FileNotFoundError) as exc:
        print(f'ERROR: {exc}',file=sys.stderr); return 2
    except KeyboardInterrupt:
        print('Interrupted. No new completed result is claimed.',file=sys.stderr); return 130
    except Exception:
        LOG.exception('Execution failed; do not treat an incomplete file as a valid result.')
        return 1


if __name__=='__main__':
    raise SystemExit(main())

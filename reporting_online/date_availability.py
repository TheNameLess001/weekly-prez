"""Observed business dates; certification remains an independent quality decision."""
from datetime import timedelta
from engine import periods

def available_bounds(dataset):
    rows=[r for r in dataset.rows if r.get('grain')!='complement'] or dataset.items
    starts=[r['start'] for r in rows if r.get('start') and r.get('end')]
    ends=[r['end'] for r in rows if r.get('start') and r.get('end')]
    return (min(starts),max(ends)) if starts else None

def available_comparisons(mode,bounds):
    low,high=bounds
    choices=[];seen=set()
    # WoW/MoM use the day after the last imported day to include its completed period.
    anchor=low
    while anchor<=high+timedelta(days=1):
        current,previous=periods(mode,anchor)
        pair=(current,previous)
        if all(low<=d<=high for p in pair for d in p) and pair not in seen:
            choices.append((anchor,current,previous));seen.add(pair)
        anchor+=timedelta(days=1)
    return choices

def default_custom(bounds):
    low,high=bounds
    length=(high-low).days+1
    if length==1:return low,high,low,high
    size=min(7,length//2)
    current_start=high-timedelta(days=size-1)
    previous_end=current_start-timedelta(days=1)
    return current_start,high,previous_end-timedelta(days=size-1),previous_end

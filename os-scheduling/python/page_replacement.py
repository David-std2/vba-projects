"""
Simulador: NRU, LRU y FIFO.
- Toma snapshot DESPUÉS de procesar cada referencia.
- NRU imprime formato detallado (Page / R / M).
- LRU y FIFO imprimen formato simple (solo Page) + Fallos.
"""

from typing import List, Tuple, Dict, Any
from collections import deque
import random

# -------------------- CONFIG --------------------
frames_count = 4
algo = "NRU"               # "NRU", "LRU", "FIFO"
frame_label_start = 1       # 0 para "Marco/Frame 0", 1 para "Marco/Frame 1"

# NRU config
# Secuencia de referencias: formato "p-R" o "p-W" o "p-W/M"
# ref_sequence = ["1-R","1-R","1-W","2-R","3-R","4-R","5-R","3-W","1-W","2-W","3-R","4-R"]
# ref_sequence = ["2-R","2-W","3-R","1-R","1-W","3-R","4-W","5-R","1-R","1-W","2-R","3-W","4-R"]
ref_sequence = ["1-R","1-R","1-W/M","2-R","3-R","4-R","5-R","3-W/M","1-W/M","2-W/M","3-R","4-R"]
# ref_sequence = ["4-R","3-R","2-W/M","1-R","5-R","2-R","3-W/M","4-R","1-R","5-R","2-R","3-R","4-W/M","1-R","3-R"]
tie_break = 'fifo'          # 'fifo' | 'index' | 'random'
reset_interval = 0          # NRU: 0 = desactivar reset; > 0 = limpiar R cada N referencias
reset_when = "after"        # 'before' | 'after' (aplica el reset antes o después del paso)

# LRU config
# ref_sequence = ["7","0","1","2","0","3","0","4","2","3","0","3","2","1","2","0","1","7","0","1"]
lru_mode = 'exact'         # 'exact' | 'aging'
aging_bits = 8             # bits del contador (ej: 8)
tick_interval = 1          # cada cuantas referencias se aplica el tick (0 desactiva)
tick_when = 'after'        # 'before' | 'after' (aplica el tick antes o después de procesar la referencia)

# FIFO config
# ref_sequence = ["3","2","1","4","3","2","15","3","2","1","4","15"]
policy = "fifo_requeue"     # "fifo_queue" | "fifo_requeue" | "second_chance"
display_queue = True        # True -> mostrar cola de referencias

# Presentation
col_w = 8
left_w = 18
debug = False               # True -> imprime decisiones internas (victima, clase, etc.)
# -------------------------------------------------

# -------------------- HELPERS --------------------
def parse_ref(ref: Any) -> Tuple[int, bool]:
    """
    Accepts:
      - int  or "3"      -> returns (3, False)
      - "3-R", "3-W/M"   -> returns (3, True/False)
    """
    if isinstance(ref, int):
        return ref, False
    s = str(ref).strip()
    if "-" in s:
        parts = s.split("-")
        p = int(parts[0])
        flags = parts[1] if len(parts) > 1 else ""
        is_write = ("W" in flags) or ("M" in flags)
        return p, is_write
    # pure numeric string
    return int(s), False

def center(text: Any, width: int) -> str:
    s = str(text)
    if len(s) >= width:
        return s[:width]
    left = (width - len(s)) // 2
    return " " * left + s + " " * (width - len(s) - left)

def nru_class(R: int, M: int) -> int:
    # clase 0..3, 0 mejor
    return 2 * R + M

# -------------------- SIMULATORS --------------------
def simulate_nru(refs: List[str], frames_count:int, reset_interval:int, reset_when:str):
    """
    NRU with tie_break global: 'fifo' | 'index' | 'random'
    snapshots: list of frames states after each reference; each frame state is dict {'page','R','M','loaded_at'}
    faults: list of booleans
    """
    frames = [{"page": None, "R":0, "M":0, "loaded_at": None} for _ in range(frames_count)]
    page_to_frame: Dict[int,int] = {}
    snapshots = []
    faults = []

    for t, ref in enumerate(refs):
        # reset 'before' if applies
        if reset_interval > 0 and reset_when == "before" and t != 0 and (t % reset_interval) == 0:
            for f in frames:
                f["R"] = 0
            if debug: print(f"[t={t}] NRU reset BEFORE R bits")

        page, is_write = parse_ref(ref)

        if page in page_to_frame:
            # HIT
            idx = page_to_frame[page]
            frames[idx]["R"] = 1
            if is_write:
                frames[idx]["M"] = 1
            faults.append(False)
            if debug:
                print(f"[t={t}] HIT page {page} in frame {idx} -> R=1 M={frames[idx]['M']}")
        else:
            # FAULT
            faults.append(True)
            free_idx = next((i for i,f in enumerate(frames) if f["page"] is None), None)
            if free_idx is None:
                # choose candidates by class (min 2*R + M)
                best_cls = None
                candidates: List[int] = []
                for i,f in enumerate(frames):
                    cls = nru_class(f["R"], f["M"])
                    if best_cls is None or cls < best_cls:
                        best_cls = cls
                        candidates = [i]
                    elif cls == best_cls:
                        candidates.append(i)

                # tie-break policies
                if len(candidates) == 1:
                    victim = candidates[0]
                else:
                    if tie_break == "index":
                        victim = min(candidates)
                    elif tie_break == "random":
                        victim = random.choice(candidates)
                    else:  # fifo (default)
                        # choose candidate with smallest loaded_at (older)
                        # loaded_at should be not None for filled frames; fallback to min index
                        best_loaded = None
                        victim = candidates[0]
                        for c in candidates:
                            ld = frames[c].get("loaded_at")
                            if best_loaded is None or (ld is not None and ld < best_loaded):
                                best_loaded = ld
                                victim = c
                if debug:
                    cand_info = ", ".join(f"{c}(cls={nru_class(frames[c]['R'],frames[c]['M'])}, ld={frames[c]['loaded_at']})" for c in candidates)
                    print(f"[t={t}] FAULT page {page} -> NRU candidates: {cand_info} -> chosen victim {victim}")

                # remove old mapping
                oldp = frames[victim]["page"]
                if oldp in page_to_frame:
                    del page_to_frame[oldp]
                free_idx = victim

            # load page into free_idx
            frames[free_idx]["page"] = page
            frames[free_idx]["R"] = 1
            frames[free_idx]["M"] = 1 if is_write else 0
            frames[free_idx]["loaded_at"] = t
            page_to_frame[page] = free_idx
            if debug:
                print(f"[t={t}] Loaded page {page} into frame {free_idx} (write={is_write})")

        # snapshot AFTER handling the ref
        snap = [{"page": f["page"], "R": f["R"], "M": f["M"], "loaded_at": f["loaded_at"]} for f in frames]
        snapshots.append(snap)

        # reset 'after' if applies
        if reset_interval > 0 and reset_when == "after" and ((t + 1) % reset_interval) == 0:
            for f in frames:
                f["R"] = 0
            if debug: print(f"[t={t}] NRU reset AFTER R bits")

    return snapshots, faults

def simulate_lru(refs: List[Any], frames_count:int, mode: str='exact',
                 aging_bits:int=8, tick_interval:int=1, tick_when:str='after'):
    """
    mode: 'exact' or 'aging'
    aging_bits: number of bits in aging counters (only for mode='aging')
    tick_interval: every N references apply the aging tick (0 disables ticks)
    tick_when: 'before' or 'after' (apply tick before or after processing the ref at that t)
    """
    snapshots = []
    faults = []

    if mode == 'exact':
        # exact LRU using last_used timestamp
        frames = [{"page": None, "last_used": None, "loaded_at": None} for _ in range(frames_count)]
        page_to_frame: Dict[int,int] = {}
        for t, ref in enumerate(refs):
            page, _ = parse_ref(ref) if "-" in str(ref) else (int(ref), False)
            if page in page_to_frame:
                idx = page_to_frame[page]
                frames[idx]["last_used"] = t
                faults.append(False)
                if debug: print(f"[t={t}] LRU HIT page {page} in frame {idx}")
            else:
                faults.append(True)
                free_idx = next((i for i,f in enumerate(frames) if f["page"] is None), None)
                if free_idx is None:
                    # victim = smallest last_used (oldest)
                    victim = 0
                    best_last = frames[0]["last_used"]
                    for i in range(1, len(frames)):
                        lu = frames[i]["last_used"]
                        if lu is None:
                            victim = i; best_last = lu; break
                        if best_last is None or (lu is not None and lu < best_last):
                            victim = i; best_last = lu
                    old = frames[victim]["page"]
                    if old in page_to_frame: del page_to_frame[old]
                    free_idx = victim
                    if debug: print(f"[t={t}] LRU FAULT page {page} -> victim frame {victim} (last_used={best_last})")
                frames[free_idx]["page"] = page
                frames[free_idx]["loaded_at"] = t
                frames[free_idx]["last_used"] = t
                page_to_frame[page] = free_idx
                if debug: print(f"[t={t}] LRU loaded page {page} into frame {free_idx}")
            snapshots.append([{"page": f["page"]} for f in frames])
        return snapshots, faults

    elif mode == 'aging':
        # Aging implementation
        max_counter = (1 << aging_bits) - 1
        frames = [{"page": None, "counter": 0, "R":0, "loaded_at": None} for _ in range(frames_count)]
        page_to_frame: Dict[int,int] = {}

        def do_tick():
            # shift right and insert R as MSB, then clear R
            for fi in range(len(frames)):
                # logical right shift
                frames[fi]["counter"] >>= 1
                frames[fi]["counter"] &= max_counter
                if frames[fi]["R"]:
                    frames[fi]["counter"] |= (1 << (aging_bits - 1))
                frames[fi]["R"] = 0

        for t, raw in enumerate(refs):
            # apply tick BEFORE if configured
            if tick_interval > 0 and tick_when == 'before' and (t % tick_interval) == 0 and t != 0:
                if debug: print(f"[t={t}] Aging tick BEFORE")
                do_tick()

            page, _ = parse_ref(raw) if "-" in str(raw) else (int(raw), False)

            if page in page_to_frame:
                idx = page_to_frame[page]
                # mark reference: next tick will shift this into counter
                frames[idx]["R"] = 1
                faults.append(False)
                if debug: print(f"[t={t}] AGING HIT page {page} in frame {idx} (counter={frames[idx]['counter']})")
            else:
                faults.append(True)
                free_idx = next((i for i,f in enumerate(frames) if f["page"] is None), None)
                if free_idx is None:
                    # choose victim = minimal counter
                    victim = 0
                    best_cnt = frames[0]["counter"]
                    for i in range(1, len(frames)):
                        cnt = frames[i]["counter"]
                        if cnt < best_cnt:
                            best_cnt = cnt; victim = i
                        elif cnt == best_cnt:
                            # tie-break: oldest loaded_at
                            ld_v = frames[victim]["loaded_at"]
                            ld_i = frames[i]["loaded_at"]
                            if ld_v is None or (ld_i is not None and ld_i < ld_v):
                                victim = i
                    old = frames[victim]["page"]
                    if old in page_to_frame: del page_to_frame[old]
                    free_idx = victim
                    if debug: print(f"[t={t}] AGING FAULT page {page} -> victim frame {victim} (cnt={best_cnt})")

                # load page into free_idx: reset counter (or set MSB?) -> we set counter = 0, mark R=1 (accessed now)
                frames[free_idx]["page"] = page
                frames[free_idx]["counter"] = 0
                frames[free_idx]["R"] = 1
                frames[free_idx]["loaded_at"] = t
                page_to_frame[page] = free_idx
                if debug: print(f"[t={t}] AGING load page {page} into frame {free_idx}")

            # snapshot AFTER handling reference (but BEFORE tick if tick_when=='after')
            snap = [{"page": f["page"], "counter": f["counter"], "R": f["R"], "loaded_at": f["loaded_at"]} for f in frames]
            snapshots.append(snap)

            # apply tick AFTER if configured
            if tick_interval > 0 and tick_when == 'after' and ((t + 1) % tick_interval) == 0:
                if debug: print(f"[t={t}] Aging tick AFTER")
                do_tick()

        return snapshots, faults

def simulate_fifo(refs: List[Any], frames_count: int, policy: str):
    """
    FIFO variants:
      - fifo: pure FIFO (hit does NOT change queue)
      - fifo_requeue: on hit move frame to end (requeue)
      - second_chance: use R bit; on hit set R=1; victim search gives second chance

    Returns: snapshots, faults, queue_log
      queue_log = list of (t, action_str, queue_snapshot)
    """
    frames = [{"page": None, "loaded_at": None, "R":0, "M":0} for _ in range(frames_count)]
    page_to_frame: Dict[int,int] = {}
    fifo_queue: List[int] = []
    snapshots, faults = [], []
    queue_log: List[Tuple[int,str,List[Any]]] = []

    def log_queue(t:int, action:str):
        # Only record if queue state or action changed (avoid duplicate consecutive entries)
        current = list(fifo_queue)
        if queue_log and queue_log[-1][2] == current and queue_log[-1][1] == action:
            return
        queue_log.append((t, action, current.copy()))

    def debug_print(t, page):
        if not debug: return
        print(f"\n[t={t}] ref={page}")
        print("  before frames:", [f["page"] for f in frames])
        print("  before fifo  :", fifo_queue)
        print("  page_to_frame:", page_to_frame)

    for t, raw in enumerate(refs):
        page = int(raw) if isinstance(raw, str) and raw.isdigit() else raw
        debug_print(t, page)

        if page in page_to_frame:
            faults.append(False)
            idx = page_to_frame[page]
            # second_chance: set R on hit
            if policy == "second_chance":
                frames[idx]["R"] = 1
                log_queue(t, f"HIT set R on frame {idx}")
                if debug: print(f"[t={t}] HIT {page} -> set R on frame {idx}")
            # fifo_requeue: move index to end
            if policy == "fifo_requeue":
                if idx in fifo_queue:
                    fifo_queue.remove(idx)
                fifo_queue.append(idx)
                log_queue(t, f"HIT requeue frame {idx}")
                if debug: print(f"[t={t}] HIT {page} -> requeued frame {idx}; fifo -> {fifo_queue}")
            else:
                if debug: print(f"[t={t}] HIT {page} (no queue change)")
        else:
            faults.append(True)
            free_idx = next((i for i,f in enumerate(frames) if f["page"] is None), None)
            if free_idx is None:
                # need a victim
                if policy == "second_chance":
                    if debug: print(f"[t={t}] second_chance: scanning for R=0 victim")
                    while True:
                        if not fifo_queue:
                            raise RuntimeError("second_chance: queue empty but no free frame")
                        front = fifo_queue.pop(0)
                        if frames[front]["R"] == 0:
                            victim = front
                            log_queue(t, f"victim frame {victim} (R=0)")
                            break
                        else:
                            frames[front]["R"] = 0
                            fifo_queue.append(front)
                            log_queue(t, f"second_chance give to frame {front}")
                            if debug: print(f"    frame {front} had R=1 -> reset and move to end -> {fifo_queue}")
                    free_idx = victim
                    if debug: print(f"[t={t}] second_chance chosen victim {victim}")
                else:
                    if not fifo_queue:
                        raise RuntimeError("FIFO: queue empty but no free frame")
                    victim = fifo_queue.pop(0)
                    free_idx = victim
                    log_queue(t, f"evict frame {victim}")
                    if debug: print(f"[t={t}] {policy}: chosen victim frame {victim} (page {frames[victim]['page']})")
                oldp = frames[free_idx]["page"]
                if oldp in page_to_frame:
                    del page_to_frame[oldp]

            # load page into free_idx
            frames[free_idx]["page"] = page
            frames[free_idx]["loaded_at"] = t
            frames[free_idx]["R"] = 1 if policy == "second_chance" else 0
            frames[free_idx]["M"] = 0
            page_to_frame[page] = free_idx
            # ensure single presence then append
            if free_idx in fifo_queue: fifo_queue.remove(free_idx)
            fifo_queue.append(free_idx)
            log_queue(t, f"load page {page} into frame {free_idx}")
            if debug: print(f"[t={t}] LOAD page {page} into frame {free_idx}; fifo -> {fifo_queue}")

        # snapshot (consistent dict)
        snap = [{"page": f["page"], "R": f.get("R",0), "M": f.get("M",0), "loaded_at": f.get("loaded_at")} for f in frames]
        snapshots.append(snap)

        if debug:
            print("  after frames:", [f["page"] for f in frames])
            print("  after fifo :", fifo_queue)

    return snapshots, faults, queue_log

# -------------------- PRINT --------------------
def print_nru_table(refs: List[str], snaps: List[List[dict]], faults: List[bool]):
    cols = len(refs)
    header = " " * left_w
    for r in refs:
        header += "|" + center(r, col_w-1)
    header += "|"
    sep = "-" * left_w
    for _ in range(cols):
        sep += "+" + "-"*(col_w-1)
    sep += "+"
    print(header)
    print(sep)
    frames_count_local = len(snaps[0])
    for fi in range(frames_count_local):
        row = center(f"Marco/Frame {frame_label_start + fi}", left_w)
        for c in range(cols):
            p = snaps[c][fi]["page"]
            cell = center(str(p), col_w-1) if p is not None else " "*(col_w-1)
            row += "|" + cell
        row += "|"
        print(row)
        rowR = " " * left_w
        for c in range(cols):
            p = snaps[c][fi]["page"]
            if p is None:
                rowR += "|" + " "*(col_w-1)
            else:
                rowR += "|" + center("R=" + str(snaps[c][fi]["R"]), col_w-1)
        rowR += "|"
        print(rowR)
        rowM = " " * left_w
        for c in range(cols):
            p = snaps[c][fi]["page"]
            if p is None:
                rowM += "|" + " "*(col_w-1)
            else:
                rowM += "|" + center("M=" + str(snaps[c][fi]["M"]), col_w-1)
        rowM += "|"
        print(rowM)
        print(sep)
    rowF = " " * left_w
    for c in range(cols):
        rowF += "|" + (center("F", col_w-1) if faults[c] else " "*(col_w-1))
    rowF += "|"
    print(rowF)

def print_simple_table(refs: List[str], snaps: List[List[dict]], faults: List[bool], show_counters=False):
    cols = len(refs)
    header = " " * left_w
    for r in refs: header += "|" + center(r, col_w-1)
    header += "|"
    sep = "-" * left_w
    for _ in range(cols): sep += "+" + "-"*(col_w-1)
    sep += "+"
    print(header); print(sep)
    frames_count_local = len(snaps[0])
    for fi in range(frames_count_local):
        row = center(f"Marco/Frame {frame_label_start + fi}", left_w)
        for c in range(cols):
            p = snaps[c][fi]["page"]
            cell = center(str(p), col_w-1) if p is not None else " "*(col_w-1)
            row += "|" + cell
        row += "|"; print(row)
        # if show_counters, print counters/R next line
        if show_counters:
            rowC = " " * left_w
            for c in range(cols):
                cell = ""
                if snaps[c][fi].get("counter") is not None:
                    cell = f"c={snaps[c][fi]['counter']}"
                elif snaps[c][fi].get("R") is not None:
                    cell = f"R={snaps[c][fi].get('R',0)}"
                rowC += "|" + center(cell, col_w-1)
            rowC += "|"; print(rowC)
        print(sep)
    rowF = " " * left_w
    for c in range(cols):
        rowF += "|" + (center("F", col_w-1) if faults[c] else " "*(col_w-1))
    rowF += "|"; print(rowF)

# -------------------- RUN --------------------
if __name__ == "__main__":
    if algo == "NRU":
        snaps, faults = simulate_nru(ref_sequence, frames_count, reset_interval, reset_when)
        print(f"ALGORITHM = NRU  reset_interval = {reset_interval}  reset_when = '{reset_when}'  tie_break = '{tie_break}'")
        print_nru_table(ref_sequence, snaps, faults)
    elif algo == "LRU":
        if lru_mode == 'exact':
            snaps, faults = simulate_lru(ref_sequence, frames_count, mode='exact')
            print(f"ALGORITHM = LRU (exact)")
            print_simple_table(ref_sequence, snaps, faults)
        else:
            snaps, faults = simulate_lru(ref_sequence, frames_count, mode='aging', aging_bits=aging_bits, tick_interval=tick_interval, tick_when=tick_when)
            print(f"ALGORITHM = LRU (aging)  aging_bits = {aging_bits}  tick_interval = {tick_interval}  tick_when = '{tick_when}'")
            print_simple_table(ref_sequence, snaps, faults, show_counters=debug)
    elif algo == "FIFO":
        snaps, faults, queue_log = simulate_fifo(ref_sequence, frames_count, policy)
        print(f"ALGORITHM = FIFO (policy = {policy})")
        print_simple_table(ref_sequence, snaps, faults)
        if display_queue:
            # print the queue_log
            print("\nCambios en la cola de páginas referenciadas:")
            for t, action, q in queue_log:
                print(f" t={t:2d} | {action:<30} | cola (frames indices) = {q}")
    else:
        raise ValueError("Algoritmo desconocido: usa 'NRU', 'LRU' o 'FIFO'")

    if debug:
        print("\nDEBUG: snapshots (raw):")
        for t, s in enumerate(snaps):
            print(t, s, "F" if faults[t] else "-")
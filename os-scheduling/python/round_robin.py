processes = [
    {"name":"P1","arrival":4,"burst":5},
    {"name":"P2","arrival":1,"burst":8},
    {"name":"P3","arrival":6,"burst":8},
    {"name":"P4","arrival":2,"burst":2},
]
Q = 4

names = [p["name"] for p in processes]
arrival = {p["name"]: p["arrival"] for p in processes}
rem = {p["name"]: p["burst"] for p in processes}
finished = {p["name"]: False for p in processes}
finish_time = {}

t = 0
queue = []
cpu = None
cpu_quantum_left = 0
to_requeue_next = []
table = []

while True:
    if all(finished.values()): break

    for p in processes:
        if arrival[p["name"]] == t: queue.append(p["name"])

    if to_requeue_next:
        queue.extend(to_requeue_next)
        to_requeue_next = []

    if cpu is None and queue:
        cpu = queue.pop(0)
        cpu_quantum_left = min(Q, rem[cpu])

    symbols = {name: ' ' for name in names}
    if cpu is not None: symbols[cpu] = '0'
    for pos, name in enumerate(queue, start=1): symbols[name] = str(pos)
    table.append(symbols)

    if cpu is not None:
        rem[cpu] -= 1
        cpu_quantum_left -= 1
        if rem[cpu] == 0:
            finished[cpu] = True
            finish_time[cpu] = t + 1
            cpu = None
            cpu_quantum_left = 0
        elif cpu_quantum_left == 0:
            to_requeue_next.append(cpu)
            cpu = None

    t += 1
    if t > 1000: break

last = max(finish_time.values())
display = {name: [] for name in names}
for tc in range(0, last + 1):
    for name in names:
        if name in finish_time and finish_time[name] == tc: display[name].append('Fi')
        else:
            if tc < len(table):  display[name].append(table[tc][name])
            else: display[name].append('.')

print("t:", list(range(0, last + 1)))
for name in names:
    print(name, display[name])
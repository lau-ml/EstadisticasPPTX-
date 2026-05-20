cd "/home/lautaro/Escritorio/Telesalud HCANK/Estadísticas/Estadísticas/" && python3 - <<'EOF'
import os
import csv
from datetime import datetime

meet_dir = "meet_history/"
output_file = "consolidado.csv"

def fmt_dt(dt_str):
    dt_str = dt_str.strip().strip('"').strip("'").strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            dt = datetime.strptime(dt_str, fmt)
            return dt.strftime('%Y/%m/%d'), dt.strftime('%H:%M')
        except ValueError:
            pass
    return dt_str, ''

rows = []

for filename in sorted(os.listdir(meet_dir)):
    if not filename.endswith('.csv'):
        continue

    meeting_code = filename[:-4]
    filepath = os.path.join(meet_dir, filename)

    with open(filepath, 'r', encoding='utf-8-sig') as f:
        content = f.read()

    created_raw = ''
    ended_raw = ''
    participants = []

    lines = content.splitlines()
    in_table = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        low = stripped.lower()

        if 'created on' in low:
            parts = stripped.split('Created on', 1)[-1].strip().lstrip(',').strip()
            created_raw = parts.split(',')[0].strip()
        elif 'ended on' in low:
            parts = stripped.split('Ended on', 1)[-1].strip().lstrip(',').strip()
            ended_raw = parts.split(',')[0].strip()
        elif 'full name' in low.split(',')[0].lower():
            in_table = True
            continue
        elif in_table:
            reader = csv.reader([stripped])
            cols = next(reader)
            if len(cols) >= 3:
                name = cols[0].strip()
                first_seen_raw = cols[1].strip()
                time_in_call = cols[2].strip()
                participants.append((name, first_seen_raw, time_in_call))

    created_date, created_time = fmt_dt(created_raw)
    ended_date, ended_time = fmt_dt(ended_raw)

    for name, first_seen_raw, time_in_call in participants:
        fs_date, fs_time = fmt_dt(first_seen_raw)
        rows.append({
            'Enlace': meeting_code,
            'Fecha creación': created_date,
            'Hora creación': created_time,
            'Fecha finalización': ended_date,
            'Hora finalización': ended_time,
            'Participante': name,
            'Fecha primera aparición': fs_date,
            'Hora primera aparición': fs_time,
            'Tiempo en llamada': time_in_call,
        })

fieldnames = [
    'Enlace', 'Fecha creación', 'Hora creación',
    'Fecha finalización', 'Hora finalización',
    'Participante', 'Fecha primera aparición', 'Hora primera aparición',
    'Tiempo en llamada'
]

with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)

print(f"Archivos procesados: {len([f for f in os.listdir(meet_dir) if f.endswith('.csv')])}")
print(f"Filas escritas:      {len(rows)}")
print(f"Salida:              {output_file}")

print("\nPrimeras 3 filas:")
for r in rows[:3]:
    print(r)
EOF
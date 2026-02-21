from pathlib import Path
import re
import datetime
import xlrd
import matplotlib.pyplot as plt
import chinese_calendar as cc

try:
    import config
except Exception:
    config = None


XLS_PATH = Path('fetched_kb.xls')
OUT_PATH = Path('semester_16week_vertical.png')

COURSE_COLORS = [
    '#E8F1FF', '#EAF7EE', '#FFF4E5', '#F3ECFF', '#FFECEF',
    '#EAF4F4', '#FFF9E6', '#EAF0FF', '#F1F8EA', '#FDEEF4',
]

CHINESE_COURSES = {'数据挖掘', '神经网络与深度学习', '羽毛球'}


def parse_cell_courses(cell_info: str):
    lines = [line.strip() for line in str(cell_info).splitlines() if str(line).strip()]
    courses = []
    for idx, line in enumerate(lines):
        if '[周]' not in line or not re.search(r'\d', line):
            continue
        if idx + 2 >= len(lines):
            continue
        place = lines[idx + 1]
        section = lines[idx + 2]
        if '节' not in section:
            continue
        name_idx = idx - 2
        if name_idx >= 0 and re.fullmatch(r'\(\d+\)', lines[name_idx]):
            name_idx -= 1
        if name_idx < 0:
            continue
        name = lines[name_idx]
        courses.append({'name': name, 'week': line, 'place': place, 'section': section})
    return courses


def expand_week_numbers(week_text: str):
    raw = str(week_text).replace('，', ',').replace(' ', '')
    odd_only = '单' in raw
    even_only = '双' in raw
    raw = raw.replace('周', '')
    raw = re.sub(r'\[.*?\]', '', raw)
    raw = re.sub(r'\(.*?\)', '', raw)

    nums = []
    for item in raw.split(','):
        if not item:
            continue
        if '-' in item:
            left, right = item.split('-', 1)
            if left.isdigit() and right.isdigit():
                nums.extend(range(int(left), int(right) + 1))
        elif item.isdigit():
            nums.append(int(item))

    nums = sorted(set(nums))
    if odd_only:
        nums = [n for n in nums if n % 2 == 1]
    elif even_only:
        nums = [n for n in nums if n % 2 == 0]
    return nums


def section_sort_key(section_label: str):
    nums = [int(x) for x in re.findall(r'\d+', section_label)]
    return nums if nums else [999]


def normalize_section(section_text: str):
    s = str(section_text).replace('节', '').strip()
    s = s.strip('[]')
    return s


def split_section_slots(section_label: str):
    nums = [int(x) for x in re.findall(r'\d+', str(section_label))]
    if not nums:
        return [str(section_label)]
    if len(nums) <= 2:
        if len(nums) == 1:
            return [f'{nums[0]:02d}']
        return [f'{nums[0]:02d}-{nums[1]:02d}']

    slots = []
    i = 0
    while i < len(nums):
        if i + 1 < len(nums):
            slots.append(f'{nums[i]:02d}-{nums[i+1]:02d}')
            i += 2
        else:
            slots.append(f'{nums[i]:02d}')
            i += 1
    return slots


def parse_all_events_from_xls(path: Path):
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_index(0)

    events = []
    for col in range(1, 6):
        for row in range(3, 17):
            cell_info = ws.cell_value(rowx=row, colx=col)
            if not isinstance(cell_info, str) or not cell_info.strip() or cell_info.strip() == ' ':
                continue
            if row > 3 and cell_info == ws.cell_value(rowx=row - 1, colx=col):
                continue

            end_row = row
            while end_row + 1 < 17 and ws.cell_value(rowx=end_row + 1, colx=col) == cell_info:
                end_row += 1

            for c in parse_cell_courses(cell_info):
                section_label = normalize_section(c['section'])
                section_slots = split_section_slots(section_label)
                lang_tag = '[中]' if c['name'] in CHINESE_COURSES else '[英]'
                for w in expand_week_numbers(c['week']):
                    if 1 <= w <= 16:
                        for slot in section_slots:
                            events.append({
                                'week': w,
                                'weekday': col,
                                'section': slot,
                                'text': f"{c['name']}{lang_tag} {c['place']}",
                            })
    return events


def build_grid(events):
    weekdays = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    sections = sorted({e['section'] for e in events}, key=section_sort_key)

    per_week = {}
    for week in range(1, 17):
        grid = {(sec, day): [] for sec in sections for day in range(1, 8)}
        for e in events:
            if e['week'] == week:
                grid[(e['section'], e['weekday'])].append(e['text'])
        per_week[week] = grid
    return weekdays, sections, per_week


def get_term_start_date():
    default_date = datetime.date(2026, 3, 2)
    if config is None:
        return default_date
    raw = getattr(config, 'term_start_date', '2026-03-02')
    try:
        return datetime.datetime.strptime(str(raw), '%Y-%m-%d').date()
    except ValueError:
        return default_date


def draw_vertical_weeks(weekdays, sections, per_week, out_path: Path):
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang SC', 'Heiti SC', 'STHeiti', 'SimHei', 'Noto Sans CJK SC']
    plt.rcParams['axes.unicode_minus'] = False

    fig_h = max(34, 2.15 * 16)
    fig, axes = plt.subplots(16, 1, figsize=(14.0, fig_h))
    fig.patch.set_facecolor('#FFFFFF')
    if hasattr(axes, 'ravel'):
        axes = axes.ravel().tolist()
    elif not isinstance(axes, list):
        axes = [axes]

    term_start = get_term_start_date()

    def extract_course_name(text: str):
        t = str(text).strip()
        if not t:
            return ''
        if ' ' in t:
            return t.rsplit(' ', 1)[0].strip()
        return t

    all_course_names = sorted({
        extract_course_name(e)
        for week in per_week.values()
        for cell in week.values()
        for e in cell
        if extract_course_name(e)
    })
    course_color_map = {
        name: COURSE_COLORS[idx % len(COURSE_COLORS)]
        for idx, name in enumerate(all_course_names)
    }

    for week_idx in range(1, 17):
        ax = axes[week_idx - 1]
        ax.axis('off')

        cell_text = []
        for sec in sections:
            row = []
            for day in range(1, 8):
                texts = sorted(set(per_week[week_idx][(sec, day)]))
                row.append(' ｜ '.join(texts))
            cell_text.append(row)

        weekday_labels = []
        day_types = []
        for day in range(1, 8):
            d = term_start + datetime.timedelta(days=(week_idx - 1) * 7 + (day - 1))
            is_weekend = d.weekday() >= 5
            is_holiday, holiday_name = cc.get_holiday_detail(d)
            holiday_label = str(holiday_name) if holiday_name else ''

            label_parts = [f'{weekdays[day - 1]} {d.strftime("%m/%d")}']
            day_type = 'normal'
            if is_holiday and d.year == 2026:
                label_parts.append(f'法定节假日({holiday_label})' if holiday_label else '法定节假日')
                day_type = 'holiday'
            elif is_weekend:
                label_parts.append('周末')
                day_type = 'weekend'

            weekday_labels.append('\n'.join(label_parts))
            day_types.append(day_type)

        table = ax.table(
            cellText=cell_text,
            rowLabels=sections,
            colLabels=weekday_labels,
            loc='center',
            cellLoc='left',
            colWidths=[0.105] * 7,
            bbox=[0.10, 0.02, 0.89, 0.96],
        )

        ax.text(
            0.055,
            0.5,
            f'第{week_idx}周',
            transform=ax.transAxes,
            rotation=90,
            va='center',
            ha='center',
            fontsize=11,
            color='#111827',
            weight='bold',
            clip_on=True,
        )
        table.auto_set_font_size(False)
        table.set_fontsize(7.2)
        table.scale(1, 1.4)

        for (r, c), cell in table.get_celld().items():
            cell.set_edgecolor('#E5E7EB')
            cell.set_linewidth(0.6)
            if r == 0:
                if c >= 0 and c < len(day_types):
                    if day_types[c] == 'holiday':
                        cell.set_facecolor('#FEE2E2')
                    elif day_types[c] == 'weekend':
                        cell.set_facecolor('#E0F2FE')
                    else:
                        cell.set_facecolor('#F3F4F6')
                else:
                    cell.set_facecolor('#F3F4F6')
                cell.set_text_props(weight='bold', color='#111827', ha='center')
            elif c == -1:
                cell.set_facecolor('#F9FAFB')
                cell.set_text_props(weight='bold', color='#374151', ha='center')
            else:
                sec = sections[r - 1]
                day = c + 1
                texts = sorted(set(per_week[week_idx][(sec, day)]))
                names = {extract_course_name(t) for t in texts if extract_course_name(t)}
                if len(names) == 1:
                    only_name = next(iter(names))
                    cell.set_facecolor(course_color_map.get(only_name, '#FFFFFF'))
                else:
                    cell.set_facecolor('#FFFFFF')
                cell.set_text_props(color='#111827', ha='left')

    plt.tight_layout(h_pad=0.8, rect=(0.0, 0.01, 1, 1))
    fig.savefig(out_path, dpi=600, bbox_inches='tight', pad_inches=0.03)
    plt.close(fig)


def main():
    if not XLS_PATH.exists():
        raise SystemExit(f'找不到 {XLS_PATH}，请先运行 process.py 生成 fetched_kb.xls')

    events = parse_all_events_from_xls(XLS_PATH)
    weekdays, sections, per_week = build_grid(events)
    draw_vertical_weeks(weekdays, sections, per_week, OUT_PATH)
    print(f'已生成图片: {OUT_PATH}')
    print(f'事件总数: {len(events)}; 节次行数: {len(sections)}')


if __name__ == '__main__':
    main()

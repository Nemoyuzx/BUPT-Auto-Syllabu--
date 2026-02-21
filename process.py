
import xlrd
import random
import string
import re
import datetime
import math
import requests
import time
import os
import platform
import csv

try:
  import config as user_config
except Exception:
  user_config = None


def cfg(name, default):
  if user_config is None:
    return default
  return getattr(user_config, name, default)

# 高级设置
year = str(cfg('year', '2026')) # 年份(兼容旧字段)
xueqi = str(cfg('xueqi', '2025-2026-2')) # 学期数
begin_week = int(cfg('begin_week', 9)) # 开学的当周(兼容旧字段)
year_week = int(cfg('year_week', 53)) # 今年总周数(兼容旧字段)
term_start_date_str = str(cfg('term_start_date', '2026-03-02')) # 学期第1周周一
Combine_Trigger = bool(cfg('Combine_Trigger', True)) # 连着几节的课程是否合并
show_week_mapping = bool(cfg('show_week_mapping', True)) # 启动时打印周次日期映射(调试)


def resolve_term_start_date():
  try:
    return datetime.datetime.strptime(term_start_date_str, '%Y-%m-%d').date()
  except ValueError:
    pass

  try:
    old_rule_week = begin_week - 1
    d = datetime.datetime.strptime(f'{year}-W{old_rule_week}-1', '%Y-W%W-%w').date()
    print(f'\nterm_start_date 格式无效，已按旧规则回退为: {d}')
    return d
  except Exception:
    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    print(f'\nterm_start_date/旧规则都不可用，已回退到系统日期所在周一: {monday}')
    return monday


TERM_START_DATE = resolve_term_start_date()


def expand_week_numbers(week_text):
  raw = str(week_text).replace('，', ',').replace(' ', '')
  odd_only = ('单' in raw)
  even_only = ('双' in raw)
  raw = raw.replace('周', '')
  raw = re.sub(r'\[.*?\]', '', raw)
  raw = re.sub(r'\(.*?\)', '', raw)
  week_numbers = []
  for item in raw.split(','):
    if not item:
      continue
    if '-' in item:
      left, right = item.split('-', 1)
      if left.isdigit() and right.isdigit():
        week_numbers.extend(range(int(left), int(right) + 1))
    elif item.isdigit():
      week_numbers.append(int(item))
  week_numbers = sorted(set(week_numbers))
  if odd_only:
    week_numbers = [w for w in week_numbers if w % 2 == 1]
  elif even_only:
    week_numbers = [w for w in week_numbers if w % 2 == 0]
  return week_numbers


def calc_lesson_date(week_num, weekday_num):
  target = TERM_START_DATE + datetime.timedelta(days=(week_num - 1) * 7 + (weekday_num - 1))
  return target.strftime('%Y%m%d')


def print_week_mapping_preview():
  print('\n周次日期映射预览(学期第1周):')
  weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
  for idx, name in enumerate(weekday_names, start=1):
    date_str = calc_lesson_date(1, idx)
    print(f'  第1周{name}: {date_str}')


def parse_cell_courses(cell_info):
  """解析课表单元格文本，支持一个单元格中包含多门课(每门一般5行)。"""
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
    teacher = lines[idx - 1] if idx - 1 >= 0 else ''
    name_idx = idx - 2
    if name_idx >= 0 and re.fullmatch(r'\(\d+\)', lines[name_idx]):
      name_idx -= 1
    if name_idx < 0:
      continue
    name = lines[name_idx]
    courses.append({
      'name': name,
      'teacher': teacher,
      'week': line,
      'place': place,
      'section': section,
    })
  return courses


def write_16week_chart(rows, markdown_path='semester_16week_chart.md', csv_path='semester_16week_chart.csv'):
  weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
  chart = {(week, day): [] for week in range(1, 17) for day in range(1, 8)}
  for row in rows:
    week = row['week_num']
    day = row['weekday']
    if 1 <= week <= 16 and 1 <= day <= 7:
      chart[(week, day)].append(f"{row['name']}@{row['place']} {row['time_range']}")

  with open(csv_path, 'w', encoding='utf-8-sig', newline='') as fcsv:
    writer = csv.writer(fcsv)
    writer.writerow(['周次'] + weekday_names)
    for week in range(1, 17):
      line = [str(week)]
      for day in range(1, 8):
        line.append(' | '.join(sorted(set(chart[(week, day)]))))
      writer.writerow(line)

  with open(markdown_path, 'w', encoding='utf-8') as fmd:
    fmd.write('# 16周课表图表\n\n')
    fmd.write('| 周次 | ' + ' | '.join(weekday_names) + ' |\n')
    fmd.write('| ' + ' | '.join(['---'] * (len(weekday_names) + 1)) + ' |\n')
    for week in range(1, 17):
      cells = []
      for day in range(1, 8):
        cells.append('<br>'.join(sorted(set(chart[(week, day)]))))
      fmd.write(f"| {week} | " + ' | '.join(cells) + ' |\n')

class ProcessBar(object):
  def __init__(self, total):  # 初始化传入总数
    self.shape = ['▏', '▎', '▍', '▋', '▊', '▉']
    self.shape_num = len(self.shape)
    self.row_num = 30
    self.now = 0
    self.total = total
  def print_next(self, now=-1):   # 默认+1
    if now == -1:
      self.now += 1
    else:
      self.now = now

    rate = math.ceil((self.now / self.total) * (self.row_num * self.shape_num))
    head = rate // self.shape_num
    tail = rate % self.shape_num
    info = self.shape[-1] * head
    if tail != 0:
      info += self.shape[tail-1]
    full_info = '[%s%s] [%.2f%%]' % (info, (self.row_num-len(info)) * ' ', 100 * self.now / self.total)
    print("\r", end='', flush=True)
    print(full_info, end='', flush=True)
    if self.now == self.total:
      print('')

if platform.system().lower() == 'windows':
  os.system("cls")
else:
  os.system("clear")


print('这是一个从BUPT教务爬取课程表并转为苹果日历的脚本 BY LAWTED')
time.sleep(1)

print('---------------GIVE ME A STAR IF U LIKE!--------------')
print('Github: www.github.com/LAWTED')
print('Github: www.github.com/Lawted')
time.sleep(1)
print("------------------NOW LET'S BEGIN!!!------------------")
time.sleep(1)
BUPT_ID = str(cfg('account', '')).strip()
BUPT_PASS = str(cfg('password', '')).strip()
if BUPT_ID and BUPT_PASS:
  print('已从 config.py 读取账号密码。')
else:
  print('config.py 未提供完整账号密码，改为手动输入。')
  BUPT_ID = input('请输入你的学号: ').strip()
  BUPT_PASS = input('请输入你的新教务密码: ').strip()




# 别动
keyStr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=' # DO NOT CHANGE!!!


pb = ProcessBar(10000)


if not begin_week:
  print('---请设置开学当周在今年的第多少周,苹果日历中查看---')
  quit()
if not year_week:
  print('---请设置今年总周数---')
  quit()
if not BUPT_ID:
  print('---请填入你的学号---')
  quit()
if not BUPT_PASS:
  print('---请填入你的密码，本项目在源码公开，不存在泄露密码操作---')
  quit()
if not keyStr:
  print('---叫你别改那个---')
  quit()

print(f'学期第1周周一日期: {TERM_START_DATE}')
if show_week_mapping:
  print_week_mapping_preview()


if platform.system().lower() == 'windows':
  os.system("cls")
else:
  os.system("clear")
for i in range(1000):
  pb.print_next()

def encodeInp(input):
  output = ''
  chr1, chr2, chr3 = '', '', ''
  enc1, enc2, enc3 = '', '', ''
  i = 0
  while True:
    chr1 = ord(input[i])
    i += 1
    chr2 = ord(input[i]) if i < len(input) else 0
    i += 1
    chr3 = ord(input[i]) if i < len(input) else 0
    i += 1
    enc1 = chr1 >> 2
    enc2 = ((chr1 & 3) << 4) | (chr2 >> 4)
    enc3 = ((chr2 & 15) << 2) | (chr3 >> 6)
    enc4 = chr3 & 63
    # print(chr1, chr2, chr3)
    if chr2 == 0:
      enc3 = enc4 = 64
    elif chr3 == 0:
      enc4 = 64
    output = output + keyStr[enc1] + keyStr[enc2] + keyStr[enc3] + keyStr[enc4]
    chr1 = chr2 = chr3 = ''
    enc1 = enc2 = enc3 = enc4 = ''
    if i >= len(input):
      break
  return output
encoded = encodeInp(BUPT_ID) + '%%%' + encodeInp(BUPT_PASS)

for i in range(1000,2000):
  pb.print_next()

session = requests.session()

for i in range(2000,3500):
  pb.print_next()

l1 = session.get('https://jwgl.bupt.edu.cn/jsxsd/')
cookies1 = l1.cookies.items()
cookie = ''
for name, value in cookies1:
  cookie += '{0}={1}; '.format(name, value)



# 第二次请求，发送cookie和密码
# print(cookie)
headers = {
  'Host': 'jwgl.bupt.edu.cn',
  'Referer': 'https://jwgl.bupt.edu.cn/jsxsd/xk/LoginToXk?method=exit&tktime=1631723647000',
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36',
  "cookie": cookie
}
payload = {'userAccount': BUPT_ID, 'userPassWord': '', 'encoded': encoded}
l5 = session.post('https://jwgl.bupt.edu.cn/jsxsd/xk/LoginToXk', data=payload, headers=headers)
for i in range(3500,4600):
  pb.print_next()


for i in range(4600,5500):
  pb.print_next()

# 第三次请求带上cookie
time.sleep(3)
data = {'xnxq01id': xueqi, 'zc': '', 'kbjcmsid': '9475847A3F3033D1E05377B5030AA94D'}
url = 'https://jwgl.bupt.edu.cn/jsxsd/xskb/xskb_print.do?xnxq01id={}&zc=&kbjcmsid=9475847A3F3033D1E05377B5030AA94D'.format(xueqi)
p = session.post(url, data=data)
if 'html' in p.text:
  print('\n------------------密码错误------------------')
  quit()
f=open('a.xls','wb')
f.write(p.content)
f.close()
with open('fetched_kb.xls', 'wb') as debug_f:
  debug_f.write(p.content)
print(f'抓取课表文件大小: {len(p.content)} bytes')
wb = xlrd.open_workbook("./a.xls")
ws = wb.sheet_by_index(0)
nrows = ws.nrows
ncols = ws.ncols
for i in range(5500,7000):
  pb.print_next()

col_values = ws.col_values(colx=0)
realtime_list = col_values[3:17]
realtime_list = [time.split('\n')[1] for time in realtime_list]
all_lesson = []
for col in range(1,6):
  for row in range(3,17):
    cell_info = ws.cell_value(rowx=row, colx=col)
    if not isinstance(cell_info, str) or not cell_info.strip() or cell_info.strip() == ' ':
      continue
    if Combine_Trigger and row > 3 and cell_info == ws.cell_value(rowx=row-1, colx=col):
      continue

    end_row = row
    if Combine_Trigger:
      while end_row + 1 < 17 and ws.cell_value(rowx=end_row+1, colx=col) == cell_info:
        end_row += 1

    start_time = realtime_list[row-3].split('-')[0]
    end_time = realtime_list[end_row-3].split('-')[1]
    time_range = start_time + '-' + end_time

    for course in parse_cell_courses(cell_info):
      all_lesson.append({
        'name': course['name'],
        'teacher': course['teacher'],
        'week': course['week'],
        'place': course['place'],
        'section': course['section'],
        'time': time_range + '+' + str(col)
      })

print(f'解析到课程片段: {len(all_lesson)}')

# 写入头文件
for i in range(7000,8000):
  pb.print_next()

f=open('calendar.ics','w',encoding='utf-8')
head=['BEGIN:VCALENDAR',
'VERSION:2.0',
]
for i in head:
  f.write(i)
  f.write('\n')
def randomUID():
  return ''.join(random.sample(['z','y','x','w','v','u','t','s','r','q','p','o','n','m','l','k','j','i','h','g','f','e','d','c','b','a'], 15))

res_txt = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\n'
for i in range(8000,9900):
  pb.print_next()

def write_file(f, name, place, date, time_start, time_end):
  global res_txt
  f.write('BEGIN:VEVENT')
  f.write('\n')
  res_txt += 'BEGIN:VEVENT'
  res_txt += '\r\n'
  f.write('DTSTAMP:20201012T104622Z')
  f.write('\n')
  res_txt += 'DTSTAMP:20201012T104622Z'
  res_txt += '\r\n'
  f.write('UID:' + randomUID())
  f.write('\n')
  res_txt += 'UID:' + randomUID()
  res_txt += '\r\n'
  f.write('SUMMARY:{} {}'.format(name,place))
  f.write('\n')
  res_txt += 'SUMMARY:{} {}'.format(name,place)
  res_txt += '\r\n'
  f.write('DTSTART;TZID=Asia/Shanghai:{}T{}'.format(date, time_start))
  f.write('\n')
  res_txt += 'DTSTART;TZID=Asia/Shanghai:{}T{}'.format(date, time_start)
  res_txt += '\r\n'
  f.write('DTEND;TZID=Asia/Shanghai:{}T{}'.format(date, time_end))
  f.write('\n')
  res_txt += 'DTEND;TZID=Asia/Shanghai:{}T{}'.format(date, time_end)
  res_txt += '\r\n'
  f.write('BEGIN:VALARM')
  f.write('\n')
  res_txt += 'BEGIN:VALARM'
  res_txt += '\r\n'
  f.write('X-WR-ALARMUID:F03864BD-41F4-40EC-BF20-1E4E7930ED92')
  f.write('\n')
  res_txt += 'X-WR-ALARMUID:F03864BD-41F4-40EC-BF20-1E4E7930ED92'
  res_txt += '\r\n'
  f.write('UID:' + randomUID())
  f.write('\n')
  res_txt += 'UID:' + randomUID()
  res_txt += '\r\n'
  f.write('TRIGGER:-PT5M')
  f.write('\n')
  res_txt += 'TRIGGER:-PT5M'
  res_txt += '\r\n'
  f.write('ATTACH;VALUE=URI:Chord')
  f.write('\n')
  res_txt += 'ATTACH;VALUE=URI:Chord'
  res_txt += '\r\n'
  f.write('ACTION:AUDIO')
  f.write('\n')
  res_txt += 'ACTION:AUDIO'
  res_txt += '\r\n'
  f.write('END:VALARM')
  f.write('\n')
  res_txt += 'END:VALARM'
  res_txt += '\r\n'
  f.write('END:VEVENT')
  f.write('\n')
  res_txt += 'END:VEVENT'
  res_txt += '\r\n'
  # print(res_txt)

event_rows_for_chart = []

for l in all_lesson:
  # print(l['name'])
  name = l['name']
  week = l['week']
  week_numbers = expand_week_numbers(week)
  place = l['place']
  time = l['time']
  time_all = time.split('+')[0]
  time_start = ''.join(time_all.split('-')[0].split(':')) + '00'
  time_end = ''.join(time_all.split('-')[1].split(':')) + '00'
  time_seven = int(time.split('+')[1])
  date = []
  for week_num in week_numbers:
    date = calc_lesson_date(week_num, time_seven)
    print(name, place, date, time_start, time_end)
    write_file(f, name, place, date, time_start, time_end)
    event_rows_for_chart.append({
      'week_num': week_num,
      'weekday': time_seven,
      'name': name,
      'place': place,
      'time_range': time_all
    })

f.write('END:VCALENDAR')
res_txt += 'END:VCALENDAR'
f.close()

write_16week_chart(event_rows_for_chart)

os.remove("./a.xls")

import urllib.parse
for i in range(9900,10000):
  pb.print_next()

res_txt = urllib.parse.quote(res_txt, safe='~@#$&()*!+=:;,.?/\'')
#使用二进制格式保存转码后的文本
res_txt = 'data:text/calendar,' + res_txt
f=open('direct.txt','w',encoding='utf-8')
f.write(res_txt)
f.close()
print('已生成: semester_16week_chart.md / semester_16week_chart.csv')
print('--------------------DONE--------------------')
print('请查看README了解如何导入和使用')

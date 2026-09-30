#!/bin/bash
# The sample documents in the website and README screenshots (the coffee festival), made with the writer engine itself.
# usage: make-samples.sh <workspace folder>        (WRITER overrides the engine command; the content is fictional)
set -euo pipefail
WRITER=${WRITER:-"dotnet $(cd "$(dirname "$0")/../.." && pwd)/src/Writer.Cli/bin/Debug/net10.0/writer.dll"}
OUT=${1:?workspace folder}
mkdir -p "$OUT"; cd "$OUT"
q() { $WRITER "$@" >/dev/null; }

# ---- 1. Word: the event plan (the hero shot; mock-model.mjs moves its dates a week on) ----
F=咖啡节活动方案.docx; rm -f $F; q create $F
q set $F / --prop title="城市咖啡节活动方案" --prop page=A4
q add $F /body --type heading --prop text="城市咖啡节活动方案" --prop level=1
q add $F /body --type paragraph --prop text="本方案用于十月第二个周末在滨江公园举办的城市咖啡节，涵盖场地、日程、分工和预算，供筹备组讨论后定稿。"
q add $F /body --type heading --prop text="活动概况" --prop level=2
q add $F /body --type paragraph --prop md="活动为期两天，设置 **24 个摊位**、2 个工作坊区和 1 个比赛舞台。现场使用可回收杯具，鼓励观众自带杯子。"
q add $F /body --type heading --prop text="日程安排" --prop level=2
q add $F /body --type table --prop style=TableGrid --prop width=100% \
  --prop data='[["时间","内容","负责"],["10 月 10 日 09:00","开幕，市集开放","运营组"],["10 月 10 日 14:00","手冲工作坊","内容组"],["10 月 11 日 10:00","城市手冲比赛","赛事组"],["10 月 11 日 17:00","闭幕与颁奖","运营组"]]'
i=0; for h in 时间 内容 负责; do i=$((i+1)); q set $F "/body/table[1]/row[1]/cell[$i]" --prop md="**$h**" --prop fill=EDF3EF; done
q add $F /body --type heading --prop text="活动亮点" --prop level=2
q add $F /body --type paragraph --prop text="今年邀请了本地十余家独立咖啡馆参展，烘焙师会在现场讲解咖啡豆的产地和处理法；另外还设置了儿童涂鸦区和宠物友好区，方便一家人一起来逛。"
q add $F /body --type heading --prop text="注意事项" --prop level=2
q add $F /body --type paragraph --prop list=bullet --prop text="所有摊位需在 10 月 9 日 18:00 前完成布置。"
q add $F /body --type paragraph --prop list=bullet --prop text="现场用电由场地方统一提供，请勿自带大功率设备。"
q add $F /body --type paragraph --prop list=bullet --prop text="如遇雨天，工作坊移至室内展厅，比赛照常进行。"

# ---- 2. Excel: the budget, formulas and a chart under the table (the right side is the format or AI panel) ----
F=咖啡节预算.xlsx; rm -f $F; q create $F
S='/sheet[1]'
q set $F $S --prop name=预算
q set $F "$S/range[A1:E1]" --prop values='[["项目","预算（元）","已支出（元）","剩余（元）","进度"]]'
q set $F "$S/range[A2:C7]" --prop values='[["场地与搭建",48000,41500],["摊位物料",26000,18200],["宣传推广",15000,9800],["工作坊耗材",8000,3600],["比赛奖品",12000,12000],["安保与保洁",10000,4000]]'
q set $F "$S/cell[A8]" --prop value=合计
for r in 2 3 4 5 6 7; do q set $F "$S/cell[D$r]" --prop formula="B$r-C$r"; q set $F "$S/cell[E$r]" --prop formula="C$r/B$r"; done
for c in B C D; do q set $F "$S/cell[${c}8]" --prop formula="SUM(${c}2:${c}7)"; done
q set $F "$S/cell[E8]" --prop formula="C8/B8"
q set $F "$S/range[B2:D8]" --prop format='#,##0'
q set $F "$S/range[E2:E8]" --prop format='0%'
q set $F "$S/range[A1:E1]" --prop bold=true --prop fill=E8F0EB --prop color=1F4D36
q set $F "$S/range[B1:E8]" --prop align=right
q set $F "$S/range[A8:E8]" --prop bold=true --prop borders='{"top":"thin"}'
q set $F $S --prop widths='{"A":13,"B":11,"C":12,"D":11,"E":8}' --prop freeze=A2
q add $F $S --type chart --prop type=column --prop title="预算与已支出" --prop categories=A2:A7 \
  --prop series='[{"name":"B1","values":"B2:B7"},{"name":"C1","values":"C2:C7"}]' \
  --prop legend=bottom --prop x=0.05cm --prop y=4.6cm --prop w=12.2cm --prop h=5.5cm

# ---- 3. PowerPoint: the festival deck ----
F=咖啡节方案.pptx; rm -f $F; q create $F
box() { # slide text x y w h size color
  q add $F "//slide[$1]" --type shape --prop md="$2" --prop x=${3}cm --prop y=${4}cm --prop w=${5}cm --prop h=${6}cm --prop size=${7}pt --prop color=$8 --prop fill=none --prop line=none
}
dot() { # slide geometry x y w h fill [line]
  q add $F "//slide[$1]" --type shape --prop geometry=$2 --prop x=${3}cm --prop y=${4}cm --prop w=${5}cm --prop h=${6}cm --prop fill=$7 --prop line=${8:-none}
}
# 1 title
q add $F / --type slide --prop layout=Blank --prop background=2B2320
dot 1 ellipse 19.6 1.9 15.4 15.4 D9822B
dot 1 ellipse 17.3 11.0 4.4 4.4 F3E9DC
box 1 "**城市咖啡节**" 2.6 6.2 16 3.2 54 F3E9DC
box 1 "活动方案 · 10 月 10 日至 11 日 · 滨江公园" 2.6 9.8 18 1.6 18 D9CBB8
# 2 the three areas
q add $F / --type slide --prop layout=Blank --prop background=F7F1E8
box 2 "**三个活动区**" 2.4 1.5 20 2.2 32 2B2320
i=0
for a in "咖啡市集|本地十余家咖啡馆，每家一个摊位，杯子自带。" "手冲工作坊|烘焙师现场讲豆子的产地和处理法，每场 20 人。" "城市手冲比赛|三家合作咖啡馆各派一位评审，周日下午决出前三。"; do
  x=$(echo "2.4 + $i * 9.9" | bc); i=$((i+1))
  dot 2 roundRect $x 4.8 9.1 11.4 FFFFFF EADFD2
  dot 2 ellipse $(echo "$x + 1.0" | bc) 6.0 1.4 1.4 D9822B
  box 2 "**${a%%|*}**" $(echo "$x + 1.0" | bc) 7.8 7.4 1.8 22 2B2320
  box 2 "${a##*|}" $(echo "$x + 1.0" | bc) 9.9 7.2 4.4 14 6B5A4C
done
# 3 the schedule
q add $F / --type slide --prop layout=Blank --prop background=F7F1E8
box 3 "**日程**" 2.4 1.5 20 2.2 32 2B2320
q add $F "//slide[3]" --type table --prop x=2.4cm --prop y=4.6cm --prop w=29cm --prop h=9cm \
  --prop data='[["时间","内容","负责"],["10 月 10 日 09:00","开幕，市集开放","运营组"],["10 月 10 日 14:00","手冲工作坊","内容组"],["10 月 11 日 10:00","城市手冲比赛","赛事组"],["10 月 11 日 17:00","闭幕与颁奖","运营组"]]'
# 4 closing
q add $F / --type slide --prop layout=Blank --prop background=2B2320
box 4 "**十月见**" 2.6 7.8 20 3 44 F3E9DC

# ---- 4. Mind map: the preparation ----
F=咖啡节筹备.mm; rm -f $F; q create $F
q set $F '/topic[1]' --prop text="咖啡节筹备"
b=0
branch() { # name side children...
  local name=$1 side=$2; shift 2; b=$((b+1))
  q add $F '/topic[1]' --type topic --prop text="$name" --prop side=$side
  for c in "$@"; do q add $F "/topic[1]/topic[$b]" --type topic --prop text="$c"; done
}
branch "场地" right "搭建与布展" "用电与照明" "雨天预案"
branch "内容" right "咖啡市集" "手冲工作坊" "城市手冲比赛"
branch "宣传" left "海报与物料" "社交媒体" "合作咖啡馆"
branch "人员" left "志愿者排班" "安保与保洁"

# ---- 5. Markdown: the meeting notes ----
F=筹备会议纪要.md; rm -f $F; q create $F
q set $F /body --raw "# 咖啡节筹备会议纪要

**时间**：9 月 22 日 14:00

**参会**：运营组、内容组、赛事组

## 已确定

- 摊位上限 24 个，10 月 1 日截止报名
- 工作坊每场 20 人，线上预约
- 比赛评审由三家合作咖啡馆各派一人

## 待办

1. 运营组：确认用电方案，周五前回复场地方
2. 内容组：整理工作坊讲师名单
3. 赛事组：起草比赛规则，下次会议讨论

> 下次会议：9 月 29 日 14:00，地点不变。
"

echo samples: "$(ls "$OUT" | tr '\n' ' ')"

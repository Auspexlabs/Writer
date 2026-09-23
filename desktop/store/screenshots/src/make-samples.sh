#!/bin/bash
# The sample documents in the store screenshots, made with the writer engine itself.
# usage: make-samples.sh <workspace folder> <photo>     (the photo comes from photo.html; WRITER overrides the engine command)
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/../../../.." && pwd)
WRITER=${WRITER:-"dotnet $ROOT/src/Writer.Cli/bin/Debug/net10.0/writer.dll"}
OUT=${1:?workspace folder}; PHOTO=${2:?photo}
mkdir -p "$OUT"; cd "$OUT"
q() { $WRITER "$@" >/dev/null; }

# ---- 1. Word: a launch plan (headings, a table, lists) ----
F=秋季新品上市方案.docx; rm -f $F; q create $F
q set $F / --prop title="秋季新品上市方案" --prop page=A4 --prop margin=narrow
q add $F /body --type heading --prop text="秋季新品上市方案" --prop level=1
q add $F /body --type paragraph --prop md="本方案用于 10 月上市的**桂花拿铁**和**栗子燕麦拿铁**两款秋季限定饮品，涵盖上市目标、节奏和预算，供市场部与门店运营部讨论后定稿。"
q add $F /body --type heading --prop text="一、上市目标" --prop level=2
q add $F /body --type paragraph --prop list=bullet --prop md="首月两款新品合计售出 **12 万杯**，占门店饮品销量的 18%。"
q add $F /body --type paragraph --prop list=bullet --prop md="会员复购率提升到 **35%**，新增会员 2 万人。"
q add $F /body --type paragraph --prop list=bullet --prop md="新品话题在社交平台累计曝光 **500 万次**。"
q add $F /body --type heading --prop text="二、上市节奏" --prop level=2
q add $F /body --type table --prop style=TableGrid --prop width=100% \
  --prop data='[["阶段","时间","主要工作","负责"],["预热","9 月 22 日—30 日","海报上墙、会员预告、种草笔记","市场部"],["首发","10 月 1 日—7 日","全国门店同步上市，买一赠一","门店运营部"],["延续","10 月 8 日—31 日","第二杯半价、联名周边","市场部"],["复盘","11 月 3 日","销量与口碑复盘","数据组"]]'
i=0; for h in 阶段 时间 主要工作 负责; do i=$((i+1)); q set $F "/body/table[1]/row[1]/cell[$i]" --prop md="**$h**" --prop fill=EDF3EF; done
q add $F /body --type heading --prop text="三、预算" --prop level=2
q add $F /body --type paragraph --prop md="总预算 **68 万元**：线上投放 32 万元，门店物料 21 万元，联名周边 15 万元。"

# ---- 2. Word: a weekly report the assistant tidies up (mock-model.mjs has its edits) ----
F=市场部周报.docx; rm -f $F; q create $F
q set $F / --prop title="市场部周报" --prop page=A4
q add $F /body --type heading --prop text="市场部周报 · 9 月第 4 周" --prop level=1
q add $F /body --type paragraph --prop text="2026 年 9 月 21 日—25 日　撰写：林晓"
q add $F /body --type heading --prop text="本周进展" --prop level=2
q add $F /body --type paragraph --prop text="这周主要在忙秋季新品的预热。海报已经在 120 家门店上墙，种草笔记一共发了 36 篇，累计互动大约 8.6 万次，数据比预期好一些；另外会员预告短信周三发出，覆盖 42 万会员，打开率 12% 左右。华南区有 18 家门店的首批物料晚到两天。"
q add $F /body --type heading --prop text="下周计划" --prop level=2

# ---- 3. Excel: monthly sales with formulas and a chart ----
F=门店销售统计.xlsx; rm -f $F; q create $F
S='/sheet[1]'
q set $F $S --prop name=2026年
q set $F "$S/range[A1:G1]" --prop values='[["月份","咖啡","茶饮","烘焙","周边","合计","环比"]]'
q set $F "$S/range[A2:E10]" --prop values='[["1 月",182,96,58,21],["2 月",165,88,52,34],["3 月",198,105,61,19],["4 月",214,118,66,22],["5 月",236,131,70,25],["6 月",251,146,72,28],["7 月",268,162,69,31],["8 月",275,170,71,36],["9 月",289,158,77,42]]'
q set $F "$S/cell[A11]" --prop value=合计
for r in 2 3 4 5 6 7 8 9 10; do q set $F "$S/cell[F$r]" --prop formula="SUM(B$r:E$r)"; done
for c in B C D E F; do q set $F "$S/cell[${c}11]" --prop formula="SUM(${c}2:${c}10)"; done
for r in 3 4 5 6 7 8 9 10; do q set $F "$S/cell[G$r]" --prop formula="F$r/F$((r-1))-1"; done
q set $F "$S/range[B2:F11]" --prop format='#,##0'
q set $F "$S/range[G2:G11]" --prop format='0.0%'
q set $F "$S/range[A1:G1]" --prop bold=true --prop fill=E8F0EB --prop color=1F4D36
q set $F "$S/range[B1:G11]" --prop align=right
q set $F "$S/range[A11:G11]" --prop bold=true --prop borders='{"top":"thin"}'
q set $F $S --prop widths='{"A":8,"B":9,"C":9,"D":9,"E":9,"F":10,"G":9}' --prop freeze=A2
q add $F $S --type chart --prop type=column --prop stacked=true --prop title="1—9 月销售构成（万元）" --prop categories=A2:A10 \
  --prop series='[{"name":"B1","values":"B2:B10"},{"name":"C1","values":"C2:C10"},{"name":"D1","values":"D2:D10"},{"name":"E1","values":"E2:E10"}]' \
  --prop legend=bottom --prop x=14.2cm --prop y=0.5cm --prop w=11cm --prop h=8.2cm

# ---- 4. PowerPoint: a launch deck; slide 2 carries the photo the cut-out shot works on ----
F=秋季新品发布会.pptx; rm -f $F; q create $F
box() { # slide text x y w h size color
  q add $F "//slide[$1]" --type shape --prop md="$2" --prop x=${3}cm --prop y=${4}cm --prop w=${5}cm --prop h=${6}cm --prop size=${7}pt --prop color=$8 --prop fill=none --prop line=none
}
dot() { # slide geometry x y w h fill [line]
  q add $F "//slide[$1]" --type shape --prop geometry=$2 --prop x=${3}cm --prop y=${4}cm --prop w=${5}cm --prop h=${6}cm --prop fill=$7 --prop line=${8:-none}
}
# 1 title
q add $F / --type slide --prop layout=Blank --prop background=2B1F18
dot 1 ellipse 19.2 2.6 13.2 13.2 E3A63B
dot 1 ellipse 17.4 11.2 4.6 4.6 8A4B2A
dot 1 ellipse 28.6 1.4 2.2 2.2 F7F1E8
box 1 "**秋季新品发布会**" 2.4 6.0 17 3.2 54 F7F1E8
box 1 "桂花拿铁 · 栗子燕麦拿铁" 2.4 9.4 17 1.6 24 E3A63B
box 1 "2026 年 9 月　市场部" 2.4 15.4 12 1.2 14 A89584
# 2 the product, with the photo (a slide early in the deck: see README)
q add $F / --type slide --prop layout=Blank --prop background=F7F1E8
dot 2 ellipse 17.4 1.9 15.2 15.2 E3A63B
box 2 "**秋季限定**" 2.6 4.4 8 1.3 16 C98A1F
box 2 "**桂花拿铁**" 2.6 5.8 13 2.8 54 2B1F18
box 2 "金桂糖浆、双份浓缩和燕麦奶，入口是秋天的第一缕桂香。" 2.6 9.2 12 2.6 18 6B5A4C
box 2 "**¥ 26**" 2.6 12.2 8 1.8 32 C98A1F
box 2 "10 月 1 日起全国门店供应" 2.6 15.2 12 1.2 14 A89584
q add $F "//slide[2]" --type image --prop src="$PHOTO" --prop x=15.2cm --prop y=2.4cm --prop w=15.6cm --prop h=15.6cm --prop alt="桂花拿铁"
# 3 two drinks
q add $F / --type slide --prop layout=Blank --prop background=F7F1E8
box 3 "**两款秋季限定**" 2.4 1.5 20 2.2 36 2B1F18
for i in 1 2; do
  if [ $i = 1 ]; then x=2.4; c=E3A63B; n="桂花拿铁"; d="金桂糖浆、双份浓缩和燕麦奶，入口是秋天的第一缕桂香。"; p="¥ 26"
  else x=17.4; c=8A4B2A; n="栗子燕麦拿铁"; d="现炒板栗泥与燕麦奶一起打发，绵密温暖，适合降温的早晨。"; p="¥ 28"; fi
  dot 3 roundRect $x 4.6 14.1 12.6 FFFFFF EADFD2
  dot 3 ellipse $(echo "$x + 1.2" | bc) 5.8 2.4 2.4 $c
  box 3 "**$n**" $(echo "$x + 1.0" | bc) 9.0 12 1.8 28 2B1F18
  box 3 "$d" $(echo "$x + 1.0" | bc) 11.0 12 3.2 16 6B5A4C
  box 3 "**$p**" $(echo "$x + 1.0" | bc) 14.6 6 1.6 22 $c
done
# 4 goals
q add $F / --type slide --prop layout=Blank --prop background=FFFFFF
box 4 "**上市目标**" 2.4 1.5 20 2.2 36 2B1F18
i=0
for g in "12 万杯|首月两款新品合计销量" "35%|会员复购率" "500 万次|话题累计曝光"; do
  x=$(echo "2.4 + $i * 10" | bc); i=$((i+1))
  dot 4 rect $x 5.6 0.18 9.4 E3A63B
  box 4 "**${g%%|*}**" $(echo "$x + 0.8" | bc) 6.6 9 3 48 2B1F18
  box 4 "${g##*|}" $(echo "$x + 0.8" | bc) 10.2 9 1.6 18 7A6A5C
done
# 5 schedule
q add $F / --type slide --prop layout=Blank --prop background=F7F1E8
box 5 "**上市节奏**" 2.4 1.5 20 2.2 36 2B1F18
q add $F "//slide[5]" --type table --prop x=2.4cm --prop y=4.8cm --prop w=29cm --prop h=10cm \
  --prop data='[["阶段","时间","主要工作","负责"],["预热","9 月 22 日—30 日","海报上墙、会员预告、种草笔记","市场部"],["首发","10 月 1 日—7 日","全国门店同步上市，买一赠一","门店运营部"],["延续","10 月 8 日—31 日","第二杯半价、联名周边","市场部"],["复盘","11 月 3 日","销量与口碑复盘","数据组"]]'
# 6 closing
q add $F / --type slide --prop layout=Blank --prop background=2B1F18
dot 6 ellipse 13.9 3.4 6 6 E3A63B
box 6 "**十月，门店见**" 6.9 10.6 20 2.6 40 F7F1E8
q set $F "//slide[6]/shape[2]/paragraph[1]" --prop align=center

# ---- 5. Mind map: the launch plan ----
F=新品上市计划.mm; rm -f $F; q create $F
q set $F '/topic[1]' --prop text="秋季新品上市"
b=0
branch() { # name side children...
  local name=$1 side=$2; shift 2; b=$((b+1))
  q add $F '/topic[1]' --type topic --prop text="$name" --prop side=$side
  for c in "$@"; do q add $F "/topic[1]/topic[$b]" --type topic --prop text="$c"; done
}
branch "产品" right "桂花拿铁" "栗子燕麦拿铁" "秋季杯套与包装"
branch "推广" right "种草笔记 36 篇" "会员预告短信" "联名周边"
branch "渠道" right "全国门店" "小程序预点单" "外卖平台"
branch "节奏" left "9 月下旬预热" "10 月 1 日首发" "11 月初复盘"
branch "预算 68 万" left "线上投放 32 万" "门店物料 21 万" "联名周边 15 万"
branch "风险" left "桂花原料供应" "首发客流高峰"

# ---- 6. Markdown: a trip plan ----
F=京都四日行程.md; rm -f $F; q create $F
q set $F /body --raw "# 京都四日行程

十月下旬红叶刚转色，住在**四条河原町**，景点大多步行可达。

## 行前清单

- [x] 护照和签证
- [x] 往返机票（10 月 24 日—27 日）
- [ ] 关西周游卡

## 每日安排

| 日期 | 上午 | 下午 | 晚上 |
| --- | --- | --- | --- |
| 10/24 | 抵达关西机场 | 入住，逛锦市场 | 先斗町吃晚饭 |
| 10/25 | 伏见稻荷大社 | 东福寺看红叶 | 祇园散步 |
| 10/26 | 岚山竹林 | 嵯峨野小火车 | 鸭川边吃烤串 |
| 10/27 | 清水寺 | 二年坂、三年坂 | 返程 |
"

# ---- 7. Word: a contract with the lawyer's tracked changes and my comments ----
F=联名合作协议.docx; rm -f $F; q create $F
q set $F / --prop title="门店联名合作协议" --prop page=A4 --prop author="王律师"
q add $F /body --type heading --prop text="门店联名合作协议" --prop level=1
q add $F /body --type paragraph --prop text="甲方：市场部　　乙方：插画工作室"
q add $F /body --type heading --prop text="一、合作内容" --prop level=2
q add $F /body --type paragraph --prop md='乙方为甲方的秋季新品设计联名杯套与海报，共<del data-author="王律师">三</del><ins data-author="王律师">四</ins>款图案，于 9 月 15 日前交付全部源文件。'
q add $F /body --type heading --prop text="二、费用与支付" --prop level=2
q add $F /body --type paragraph --prop md='设计费共计人民币 80,000 元，签约后支付 50%，<del data-author="王律师">交付后</del><ins data-author="王律师">甲方验收通过后 10 个工作日内</ins>支付其余 50%。'
q add $F /body --type heading --prop text="三、知识产权" --prop level=2
q add $F /body --type paragraph --prop text="联名图案的著作权归双方共有，任何一方单独用于其他商业用途，须事先取得另一方书面同意。"
q add $F /body --type heading --prop text="四、保密" --prop level=2
q add $F /body --type paragraph --prop md='双方对合作中获知的未公开信息负有保密义务，期限至合作结束后<del data-author="王律师">一</del><ins data-author="王律师">两</ins>年。'
q add $F '/body/paragraph[3]' --type comment --prop author="林晓" --prop date="2026-09-23T06:30:00Z" --prop quote="验收通过" --prop text="验收标准要不要写进附件？"
q add $F '/body/paragraph[4]' --type comment --prop author="林晓" --prop date="2026-09-23T06:42:00Z" --prop quote="著作权归双方共有" --prop text="建议补充署名方式和使用期限。"
q set $F / --prop track=true --prop author="林晓"

echo samples: "$(ls "$OUT" | tr '\n' ' ')"

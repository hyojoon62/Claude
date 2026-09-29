import copy, sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = {i: s for i, s in enumerate(prs.slides, 1)}
U = 9144  # 0.01 inch

def sh(n, sid):
    for s in S[n].shapes:
        if s.shape_id == sid:
            return s
    raise KeyError((n, sid))

def set_para(shape, pi, *texts):
    p = shape.text_frame.paragraphs[pi]
    runs = p.runs
    assert len(runs) >= len(texts), (shape.shape_id, pi, len(runs), texts)
    for r, t in zip(runs, texts):
        r.text = t
    for r in runs[len(texts):]:
        r._r.getparent().remove(r._r)

def del_para(shape, pi):
    p = shape.text_frame.paragraphs[pi]._p
    p.getparent().remove(p)

def style_run(r, size=None, color=None, font=None):
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if font:
        rPr = r._r.get_or_add_rPr()
        for tag in ('latin', 'ea', 'cs'):
            el = rPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}' + tag)
            if el is not None: el.set('typeface', font)

def set_paras(shape, spec):
    """spec: list of (template_pi, [ (text, size, color, font) ... ])"""
    tf = shape.text_frame
    tmpl = [copy.deepcopy(p._p) for p in tf.paragraphs]
    txBody = tf._txBody
    for p in list(tf.paragraphs):
        txBody.remove(p._p)
    for tpi, runs in spec:
        newp = copy.deepcopy(tmpl[min(tpi, len(tmpl) - 1)])
        txBody.append(newp)
        from pptx.text.text import _Paragraph
        para = _Paragraph(newp, tf)
        rs = para.runs
        base = rs[-1]._r
        for r in rs[len(runs):]:
            newp.remove(r._r)
        while len(para.runs) < len(runs):
            newp.append(copy.deepcopy(base))
        for r, (t, size, color, font) in zip(para.runs, runs):
            r.text = t
            style_run(r, size, color, font)

def geom(shape, L=None, T=None, W=None, H=None):
    if L is not None: shape.left = Emu(int(L * U))
    if T is not None: shape.top = Emu(int(T * U))
    if W is not None: shape.width = Emu(int(W * U))
    if H is not None: shape.height = Emu(int(H * U))

def remove(shape):
    shape._element.getparent().remove(shape._element)

TEAL, DARK, GRAY, LIGHT = '174E58', '26282B', '5B6066', '969CA2'
SB, RG = 'Pretendard SemiBold', 'Pretendard'

# ---------------------------------------------------------------- 0. sync to latest PDF (2026-09-28 18:19)
# slide 1: 서울(출국 전) first on the timeline
emu_px = 12191695 / 800
c_seoul = 731520
c_eu0 = c_seoul + 105 * emu_px
step = 75 * emu_px
dot_r = 59436
eu = [(12, 11), (14, 13), (16, 15), (18, 17), (20, 19), (22, 21), (24, 23), (26, 25)]  # (label id, oval id)
for k, (lab, ov) in enumerate(eu):
    c = c_eu0 + k * step
    sh(1, ov).left = Emu(int(c - dot_r))
    sh(1, lab).left = Emu(int(c - 45720))
sh(1, 27).left = Emu(int(c_seoul - dot_r))
sh(1, 28).left = Emu(int(c_seoul - 45720)); sh(1, 28).width = Emu(1188720)
seoul = sh(1, 28)
set_paras(seoul, [(0, [('서울', None, None, None), (' (출국 전)', 9, LIGHT, RG)])])
solid, dashed = sh(1, 9), sh(1, 10)
solid.left = Emu(int(c_eu0)); solid.width = Emu(int(7 * step))
dashed.left = Emu(int(c_seoul)); dashed.width = Emu(int(c_eu0 - c_seoul))
set_para(sh(1, 29), 0, '서울 2곳 (10~11월, 출국 전)  ·  유럽 8개 도시 (2026.12.22 ~ 2027.1.3)')

set_para(sh(2, 22), 0, '서울 주차칸 위치 후보 지도  ·  킥보드 없는 거리 밖으로 방치가 옮겨 갔는지 평가')
set_para(sh(4, 14), 0, '서울이 한 것')
set_para(sh(4, 18), 0, '아직 없는 것')
set_para(sh(7, 16), 0, '100m 단위로 환산')
set_para(sh(8, 31), 0, '위험한 위치의 정의는 서울 기준에서 빌려 씁니다.',
         ' 현지 규정을 지켰는지는 따지지 않고, 보행자에게 위험한 위치인지만 봅니다.')
set_para(sh(10, 2), 0, '04', '   방법 · 전공 연계')
for n, suffix in [(12, ''), (13, ' · 독일'), (14, ' · 벨기에·프랑스'), (15, ' · 스페인'), (16, '')]:
    set_para(sh(n, 2), 0, '05', '   국가별 활동 세부 내용' + suffix)
set_para(sh(16, 11), 0, '한 도시의 현장 촬영은 약 1시간, 분류와 판정은 당일 저녁에 합니다')
set_para(sh(17, 21), 1, '비율과 100m당 △·× 대수를 함께 비교')
set_paras(sh(19, 21), [(0, [('사전 조사(완료)', None, None, None)]), (0, [('기준표 초안', None, None, None)])])
set_para(sh(20, 2), 0, '08', '   기대효과 · 향후 활용')
set_para(sh(20, 11), 0, '결과는 제안서로 정리해 서울시·자치구 시민 제안 창구에 제출합니다')
set_para(sh(20, 29), 1, '제안서를 서울시·자치구 시민 제안 창구에 제출, 검토 회신 요청')

# ---------------------------------------------------------------- A. 수치 수정
# slide 3
set_para(sh(3, 18), 0, '6.2만', ' 건')
set_para(sh(3, 19), 1, '2021년 약 2.1만 건에서 2.9배')
set_paras(sh(3, 31), [(0, [('공유자전거 민원도 2024년 월평균 323건으로 전년의 1.6배', None, None, None)]),
                     (0, [('(국민권익위원회, 2024.10)', None, None, None)])])
set_para(sh(3, 12), 0, '출처: 파이낸셜뉴스(2024.11, 서울시 통계) · 국민권익위원회(2024.10)')

# slide 4
set_para(sh(4, 20), 1, "생활인구 500명 설문, '무단 방치 감소' 80.4%")
s17 = sh(4, 17)
set_paras(s17, [(0, [('2026.9.11 조례 개정', None, None, None)]),
                (1, [('지정주차구역 설치·운영 근거 명문화', None, None, None)]),
                (1, [('이미 24개 자치구 485곳 운영 중', 13, GRAY, None)])])
set_para(sh(4, 21), 1, '어디에 우선 늘리고 어떤 반납 규칙을 둘지는 기준이 없음')
set_paras(sh(4, 25), [(0, [('설문은 실제와 다를 수 있습니다. 워싱턴DC·오클랜드 연구에서 시민은 20~30% 이상이 잘못 세워졌다고 느꼈지만,', None, None, None)]),
                     (0, [('실제 규정 위반은 15~19%, 보행 방해는 5~6%였습니다 (Klein 외, 2023).', None, None, None)])])
set_para(sh(4, 12), 0, '출처: 서울시(2026.9, 킥보드 없는 거리 확대·생활인구 설문) · 서울특별시의회(2026.9.11) · '
         '전국매일신문(2026.9) · Klein 외(2023)')

# slide 5
set_para(sh(5, 11), 0, '규제 전후를 같은 방법으로 센 유럽 보도 조사는 찾지 못했습니다')
set_para(sh(5, 17), 0, '규제 전후 비교 조사')
set_para(sh(5, 21), 1, '마드리드  ', '공유 킥보드 6,000대 허가 취소(2024.10)')
set_para(sh(5, 26), 0, '프랑크푸르트  ', '주차면 263곳(2026.6), 반경 100m 안 반납 금지')
set_para(sh(5, 12), 0, '출처: paris.fr · NPR(2023) · madrid.es(2024.9) · The Objective(2024.10) · hessenschau(2026.6) · '
         'Brussels Times(2024~2026) · heidelberg.de · barcelona.cat')
# footnote under the table (clone the source line style)
src = sh(5, 12)
fn = copy.deepcopy(src._element)
src._element.addnext(fn)
fnote = [s for s in S[5].shapes if s._element is fn][0]
fnote._element.nvSpPr.cNvPr.id = 90
fnote._element.nvSpPr.cNvPr.name = 'TextBox Footnote'
geom(fnote, L=75, T=604, W=1183, H=25)
set_para(fnote, 0, '보도 위 킥보드를 센 연구는 미국·뉴질랜드·노르웨이 등에 있으나 대부분 규제 이전 관찰입니다 '
         '(Brown 외 2020, Klein 외 2023, Karlsen 외 2021)')
style_run(fnote.text_frame.paragraphs[0].runs[0], size=10, color=GRAY)

# slide 9
set_para(sh(9, 35), 0, '선례  ', '공유 킥보드 업체 Lime은 2021년 텔아비브에서 AI로 주차 합법 여부를 판정해 경고하는 기능을 도입했습니다.')
set_para(sh(9, 12), 0, '출처: Globes(2021.6)')

# ---------------------------------------------------------------- C. 13~15쪽 도시별 표
# columns (0.01 in): 도시·일정 | 구분 | 규제 방식 | 장소 | 이 도시에서 보는 것
COL = dict(city=(75, 140), tag=(222, 50), rule=(282, 262), place=(556, 168), see=(736, 522))
HDR = [(14, 'city', '도시 · 일정'), (15, 'tag', '구분'), (16, 'rule', '규제 방식'), (17, 'place', '장소'),
       (18, 'see', '이 도시에서 보는 것')]

PLACE = {(13, 2): ['대성당 광장', '라인강 다리'], (14, 1): ['루브르~리볼리', '400m 구간']}

def table(n, rows, row_top, row_h, gap=10):
    for sid, col, text in HDR:
        s = sh(n, sid); L, W = COL[col]; geom(s, L=L, W=W); set_para(s, 0, text)
    for k, (ids, line_id, rule, see) in enumerate(rows):
        T = row_top + k * (row_h + gap + 1)
        city, tag, r, place, look = (sh(n, i) for i in ids)
        for s, col in zip((city, tag, r, place, look), ('city', 'tag', 'rule', 'place', 'see')):
            L, W = COL[col]; geom(s, L=L, T=T, W=W, H=row_h)
        # 규제 방식: 첫 줄 방식(semibold), 둘째 줄 조치
        det = rule[1] if isinstance(rule[1], (list, tuple)) else [rule[1]]
        set_paras(r, [(0, [(rule[0], 12.5, None, None)])] + [(1, [(d, 11.5, None, None)]) for d in det])
        if (n, k) in PLACE:
            lines = PLACE[(n, k)]
            set_paras(place, [(0, [(t, None, None, None)]) for t in lines])
        # 보는 것: 첫 줄 무엇을 세고 비교하는지(teal), 둘째 줄 목적(gray)
        spec = [(0, [(see[0], 12, TEAL, SB)])]
        if len(see) > 1:
            spec.append((0, [(see[1], 11, GRAY, RG)]))
        set_paras(look, spec)
        for p in look.text_frame.paragraphs[1:]:
            p.space_before = Pt(3)
        if line_id:
            geom(sh(n, line_id), T=T + row_h + gap / 2)

# slide 13 독일
set_para(sh(13, 11), 0, '프랑크푸르트에서는 주차면 근처와 먼 곳의 보도를 비교합니다')
table(13, [
    ((20, 21, 22, 23, 24), 25, ('관리', ['주차면 263곳', '반경 100m 안 반납 금지']),
     ('주차면 100m 안·밖 각 200m의 100m당 기기 수와 방해·위험 비율 비교',
      '규칙이 지켜지는 구간의 보도가 실제로 더 비어 있는지 확인')),
    ((26, 27, 28, 29, 30), 31, ('관리', ['3개 업체 1,200대', '구시가는 지정 구역에만 주차']),
     ('구시가 지정 구역의 간격·크기·표시 방식을 사진과 지도로 기록',
      '서울 주차칸 설계 참고')),
    ((32, 33, 34, 35, 36), 37, ('관리', '보행자구역·다리·강변 반납 금지'),
     ('금지 구역 경계를 바닥 표시·표지판·앱 중 무엇으로 알리는지 기록',)),
], row_top=255, row_h=95)
set_para(sh(13, 12), 0, '출처: hessenschau(2026.6) · frankfurt.de · heidelberg.de · stadt-koeln.de')

# slide 14 벨기에·프랑스
table(14, [
    ((20, 21, 22, 23, 24), 25, ('관리 (2027 금지 예정)', ['주차구역 1,700여 곳', '2025년 부상 666명, 12월 마지막 운행']),
     ('주차구역 근처·먼 곳 200m씩 비교해 촘촘한 관리로도 보도가 비워지는지 확인',
      '금지 후에도 남는 공유 자전거 수를 함께 기록')),
    ((26, 27, 28, 29, 30), 31, ('금지', ['2023 주민투표로 1.5만 대 퇴출', '공유 전기자전거는 운영 중']),
     ('남은 공유 전기자전거를 세고 주차 상태(○ △ ×) 분류',
      '브뤼셀(관리)·바르셀로나(불허)와 비교해 문제가 다른 기기로 옮겨 갔는지 확인')),
], row_top=255, row_h=110)
note_bar, note = sh(14, 32), sh(14, 33)
geom(note_bar, T=530); geom(note, T=530)
set_para(note, 0, '브뤼셀  ', '2027년부터 공유 킥보드는 금지되지만 공유 자전거는 계속 허용됩니다.')

# slide 15 스페인
table(15, [
    ((20, 21, 22, 23, 24), 25, ('금지', '2024.10 공유 킥보드 6,000대 허가 취소'),
     ('퇴출 2년 뒤 중심가 보도에 남은 공유 기기 종류와 거치대 위치 기록',)),
    ((26, 27, 28, 29, 30), 31, ('불허', ['2019년부터 공유 킥보드 불허', '거치대형 공공자전거만 운영']),
     ('거치대 간격과 위치를 기록',
      "'거치대만 있는 도시'의 보도 모습을 서울 비교 자료로 정리")),
    ((32, 33, 34, 35, 36), 37, ('불허', ['공유 킥보드 허가 없음', '공유 전기자전거·스쿠터는 허가']),
     ('공유 전기자전거를 세고 주차 상태(○ △ ×) 분류',
      '킥보드가 없는 도시의 기준값으로 사용')),
], row_top=255, row_h=90)
set_para(sh(15, 12), 0, '출처: madrid.es(2024.9) · The Objective(2024.10) · El Español(2024.9) · barcelona.cat')


# ---------------------------------------------------------------- 가독성: 의미 단위 줄바꿈
def lines(n, sid, *ls):
    set_paras(sh(n, sid), [(0, [(t, None, None, None)]) for t in ls])
lines(6, 22, '킥보드를 막아도 그 자리를 다른 공유 기기가 채울 수 있습니다.', '그래서 공유 전기자전거까지 함께 셉니다.')
lines(6, 25, '유럽 4개 도시와 서울 2곳의 보도 위 공유 이동수단을', '같은 기준으로 셉니다')
lines(6, 27, '규제 방식, 주차구역과의 거리, 금지 구역 안팎에 따라', '주차 상태를 비교합니다')
set_paras(sh(9, 28), [(0, [('1', None, None, None)]), (1, [('출국 전 서울 사진을', None, None, None)]), (1, [('두 사람이 따로 판정해 일치율 확인', None, None, None)])])
set_paras(sh(9, 30), [(0, [('2', None, None, None)]), (1, [('이미지를 보는 AI 모델에', None, None, None)]), (1, [('기준표와 예시 사진을 주고 분류', None, None, None)])])

cp = prs.core_properties
cp.author = ''
cp.last_modified_by = ''
prs.save(DST)
print('saved', DST)

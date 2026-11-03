# storyboard.md + storyboard.xlsx 를 같은 데이터로 만든다. 실행: python3 storyboard_build.py
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.units import pixels_to_EMU

HERE = os.path.dirname(os.path.abspath(__file__))
F = "analysis/frames/"

# (씬, 시작, 끝, 화면 설명, 자막·카피, 오디오, 참고 패턴, 참고 컷[], 메모, AI 프롬프트)
SCENES = [
    ("S01", 0, 4, "밤 야구장. 관중이 가득 찬 스탠드와 조명탑 아래, 마운드 위를 걸어가는 신랑의 뒷모습. 카메라가 등 뒤에서 자연스러운 핸드헬드로 따라감 (허리~머리 높이)",
     "없음", "음악 도입(잔잔) + 관중 함성 앰비언스", "03 c02 뒷모습 팔로우 / 01 c02 얼굴 숨긴 걷기",
     ["03/c02", "01/c02"], "첫 컷부터 얼굴 숨김. 관중 플래시가 반짝여 '무대에 선 남자' 느낌",
     "[OUTFIT], seen from behind walking across the pitcher's mound of a packed night baseball stadium, handheld follow shot from behind at shoulder height, stadium floodlights, crowd camera flashes in the stands, [STYLE]"),
    ("S02", 4, 8, "마운드에서 내려와 더그아웃 쪽으로 뚜벅뚜벅 걸어감. 계속 등 뒤 팔로잉, 앞쪽에 더그아웃과 선수들이 점점 가까워짐",
     "없음", "관중 함성 + 발소리", "01 c02 걷기로 존재감 쌓기",
     ["01/c02"], "속도는 여유 있게. 서두르지 않는 걸음",
     "[OUTFIT], seen from behind, walking steadily from the pitcher's mound toward the dugout, players in uniform waiting near the dugout ahead, handheld follow shot, night stadium lights, [STYLE]"),
    ("S03", 8, 11, "더그아웃 앞 선수들이 환호하며 손을 내밂. 신랑이 여유 있게 주먹 인사로 답함. 오버숄더(신랑 어깨 너머)로 선수들만 보이고 신랑 얼굴은 안 보임",
     "없음", "0:08 음악 빌드업 시작 + 선수들 환호", "03 c10 '배우' 같은 캐릭터 포즈",
     ["03/c10"], "음악이 커지는 0:08에 주먹 인사를 맞추면 좋음",
     "over-the-shoulder shot from behind [OUTFIT], casually fist-bumping cheering baseball players in uniform at the dugout, groom's face not visible, players smiling, night stadium, handheld, [STYLE]"),
    ("S04", 11, 13, "선수 통로(터널) 입구로 들어가는 실루엣. 뒤의 경기장 빛이 역광으로 남고 안쪽은 어두움",
     "없음", "함성이 먹먹하게 멀어짐", "01 c06 어두운 공간으로 긴장 전환",
     ["01/c06"], "여기서부터 핸드헬드 → 정돈된 컷 전환으로 리듬 바뀜",
     "silhouette of [OUTFIT] from behind entering a dark stadium players' tunnel, bright stadium light behind him as backlight, dark corridor ahead, [STYLE]"),
    ("S05", 13, 15, "걷는 구두 발끝 클로즈업. 바닥 가까이 낮은 앵글로 트래킹",
     "없음", "구두 발소리 또렷하게", "01 c01 / 02 c06 발 클로즈업 훅",
     ["01/c01", "02/c06"], "",
     "low-angle close-up tracking shot of polished black oxford shoes walking on a polished concrete corridor floor, [OUTFIT] trouser hems, [STYLE]"),
    ("S06", 15, 18, "복도 측면 풀샷. 카메라는 고정, 인물이 왼쪽에서 프레임 인 → 오른쪽으로 프레임 아웃. 벽에 구단 엠블럼·트로피 진열장",
     "없음", "빌드업", "01 c03 공간으로 무게감 주기 (측면 구도로 비틂)",
     ["01/c03"], "프레임 아웃 후 빈 복도 0.5초 남기기",
     "static wide side shot of a stadium interior corridor with trophy cabinets and team emblem on the wall, [OUTFIT] walking into frame from left and exiting right, face turned away, [STYLE]"),
    ("S07", 18, 20, "복도 정면 대칭 와이드. 긴 복도 끝에 두껍고 무거운 문, 신랑이 작게 걸어감(뒷모습)",
     "없음", "빌드업", "01 c03 대칭 와이드 + 문",
     ["01/c03", "user/u02"], "문은 실제 예식장 문처럼 가운데가 갈라지는 양문",
     "symmetrical wide shot down a long dim corridor, [OUTFIT] small in frame walking away toward a pair of tall heavy wooden double doors at the end, [STYLE]"),
    ("S08", 20, 22, "문 앞에 멈춰 선 뒷모습 미디엄샷. 문틈 아래로 희미한 빛",
     "없음", "빌드업 고조", "01 c06 '곧 이 문이 열린다'",
     ["01/c06", "user/u02"], "",
     "medium shot from behind of [OUTFIT] standing still in front of tall heavy wooden double doors, faint light leaking under the doors, [STYLE]"),
    ("S09", 22, 23.5, "커프스 단추를 만지는 손 익스트림 클로즈업",
     "없음", "빌드업 고조", "01 c04 디테일 '준비 완료'",
     ["01/c04", "user/u04"], "",
     "extreme close-up of a hand adjusting a cufflink on a white shirt cuff under a black suit sleeve, [STYLE]"),
    ("S10", 23.5, 25, "재킷 단추를 잠그는 손 클로즈업 (가슴~허리, 얼굴 X)",
     "없음", "빌드업 고조", "01 c05 / 02 c05 디테일",
     ["01/c05", "user/u05"], "",
     "close-up of hands buttoning a black suit jacket over a matching black vest, white boutonniere on the lapel, framed from chest to waist, face out of frame, [OUTFIT], [STYLE]"),
    ("S11", 25, 26.5, "보타이를 바로잡고, 손이 내려와 가슴의 흰 꽃 코사지를 살짝 매만지는 클로즈업 (턱 아래까지만)",
     "없음", "빌드업 고조", "01 c04~c05 디테일 연속",
     ["01/c05", "user/u06"], "본식 의상 확정: 보타이 + 흰 코사지. 코사지가 돋보이는 유일한 컷이라 초점을 코사지에",
     "close-up of fingers straightening a black bow tie, then lightly touching a white flower boutonniere on the black suit lapel, framed below the chin, face not visible, shallow focus on the boutonniere, [OUTFIT], [STYLE]"),
    ("S12", 26.5, 28, "양손이 문 손잡이를 동시에 잡는 클로즈업",
     "없음", "빌드업 최고조", "01 c07 화면 속 문 → 실제 문 복선",
     ["01/c07", "user/u01"], "",
     "close-up of both hands gripping the long vertical handles of heavy wooden double doors, white shirt cuffs and black suit sleeves, [STYLE]"),
    ("S13", 28, 30, "문이 살짝 열리며 문틈으로 새하얀 빛이 새어 나옴. 로우앵글 뒷모습 실루엣",
     "없음", "문 여는 무거운 소리, 0:29.5 음악 끊김", "user u02 문틈 빛 / 01 c12 문 이음새 = 실제 문",
     ["user/u02"], "빛이 갈라지는 자리 = 화면 정중앙 = 실제 문 이음새",
     "low-angle shot from behind of [OUTFIT] pushing open heavy wooden double doors, a narrow vertical gap of brilliant white light in the center, lens flare, [STYLE]"),
    ("S14", 30, 32, "양팔로 문을 활짝 밀어 열고 빛 속으로 들어감 → 빛이 화면 전체를 덮으며 화이트 디졸브",
     "없음", "0:30 음악 터짐", "user u03 / 01 c10·03 c15 화이트 플래시",
     ["user/u03", "01/c10"], "이 순간 실제 예식장 문 열림. 문 뒤 하얀 역광 조명 요청 권장",
     "[OUTFIT] from behind with both arms spread pushing the double doors wide open into overwhelming pure white light, silhouette dissolving into white, [STYLE]"),
    ("S15", 32, 33, "풀 화이트 홀드",
     "없음", "음악 계속", "03 c15 플래시 후 대기 화면 연결",
     ["03/c15"], "실제 문이 열리는 동안 화면은 하얗게 유지",
     "pure white frame, soft bloom, [STYLE]"),
    ("W1", 33, 36, "화이트가 걷히며 밤 야구장 상공이 드러남. 첫 불꽃이 양옆에서 터짐 (1회 재생)",
     "없음", "음악 계속 + 불꽃 소리", "02 c11 클라이맥스 뒤 감속 구간",
     ["02/c11"], "정중앙은 비워 둠",
     "white light fading to reveal a night baseball stadium from above the field, first gold fireworks bursting on the left and right sides, center of frame kept empty, [STYLE]"),
    ("W2", 36, 56, "대기 화면(20초 루프): 밤 경기장 위로 금색·흰색 불꽃이 좌우에서 천천히 터지고 종이 가루가 내려옴. 좌우 대칭, 정중앙 비움",
     "없음", "음악 계속", "03 c16 / 02 c12 갈라져도 예쁜 대기 화면",
     ["03/c16", "02/c12"], "처음과 끝 화면이 같게 만들어 이음매 없이 반복. 문이 열린 양쪽에 각각 그림이 보여야 함",
     "seamless looping shot of a night baseball stadium under restrained gold and white fireworks bursting symmetrically on both sides, slowly falling gold confetti, stadium floodlights, center of frame kept empty, [STYLE]"),
]

COLS = ["씬", "시간", "화면 설명", "자막·카피", "오디오", "참고 패턴 (영상 번호·컷)", "참고 컷", "메모", "AI 프롬프트"]


def t(s):
    return f"{int(s // 60)}:{s % 60:04.1f}".replace(".0", "")


def build_md():
    # 본편은 정확히 33초여야 한다 (brief)
    assert sum(x[2] - x[1] for x in SCENES if x[0].startswith("S")) == 33
    assert sum(x[2] - x[1] for x in SCENES if x[0] == "W2") == 20
    lines = [
        "# 스토리보드 — 신랑 입장 영상",
        "",
        "- 본편 **0:00~0:33 (33초)**, 문 열림 신호 = S13~S14 (0:28~0:32). 이어서 대기 화면 W1(1회) + **W2 20초 루프**",
        "- 16:9 mp4 / 소스 전부 생성형 AI / 신랑 얼굴 끝까지 숨김",
        "- AI 프롬프트의 `[OUTFIT]`, `[STYLE]`은 `brief.md`의 공통 문구로 바꿔 넣는다 (의상 사진을 받으면 `[OUTFIT]`만 수정)",
        "- 참고 컷의 `user/`는 사용자가 보낸 참고 이미지, `01·02·03/`은 레퍼런스 캡처",
        "",
        "| " + " | ".join(COLS) + " |",
        "|" + "---|" * len(COLS),
    ]
    for sc, s, e, desc, cap, audio, pat, refs, memo, prompt in SCENES:
        imgs = " ".join(f'<img src="{F}{r}.jpg" width="140">' for r in refs)
        row = [f"**{sc}**", f"{t(s)}–{t(e)}<br>({e - s:g}초)", desc, cap, audio, pat, imgs, memo, f"`{prompt}`"]
        lines.append("| " + " | ".join(row) + " |")
    open(os.path.join(HERE, "storyboard.md"), "w").write("\n".join(lines) + "\n")


def build_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "스토리보드"
    ws.append(COLS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F2A44")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = [7, 12, 40, 10, 22, 26, 24, 28, 55]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    for r, (sc, s, e, desc, cap, audio, pat, refs, memo, prompt) in enumerate(SCENES, start=2):
        ws.append([sc, f"{t(s)}–{t(e)}\n({e - s:g}초)", desc, cap, audio, pat, "", memo, prompt])
        ws.row_dimensions[r].height = max((82 * len(refs) + 8) * 0.75, 110)  # 포인트 단위
        for c in ws[r]:
            c.alignment = Alignment(vertical="top", wrap_text=True)
        for k, ref in enumerate(refs):  # 참고 컷을 G열 셀 안에 세로로 쌓는다
            img = XLImage(os.path.join(HERE, F + ref + ".jpg"))
            img.width, img.height = int(img.width * 76 / img.height), 76
            img.anchor = OneCellAnchor(_from=AnchorMarker(col=6, row=r - 1, colOff=pixels_to_EMU(4), rowOff=pixels_to_EMU(4 + 82 * k)),
                                       ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
            ws.add_image(img)
    ws.freeze_panes = "C2"
    wb.save(os.path.join(HERE, "storyboard.xlsx"))


if __name__ == "__main__":
    build_md()
    build_xlsx()
    print("ok")

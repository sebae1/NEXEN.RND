import numpy as np
import matplotlib.pyplot as plt
from traceback import format_exc
from matplotlib import rc, font_manager

class Config:
    PERIOD: str = "전체"
    OPENAI_API_KEY: str = ""
    CLAUDE_API_KEY: str = ""
    GPT_MODELS: list[str] = []
    LAST_USED_GPT_MODEL: str = ""
    CLAUDE_MODELS: list[str] = []
    LAST_USED_CLAUDE_MODEL: str = ""

    @classmethod
    def get_months(cls) -> list[int,]:
        match Config.PERIOD:
            case "전체":
                return list(range(1, 13))
            case "1Q":
                return [1, 2, 3]
            case "2Q":
                return [4, 5, 6]
            case "3Q":
                return  [7, 8, 9]
            case "4Q":
                return [10, 11, 12]
            case _:
                return [int(cls.PERIOD[:-1]),]

class ExceptionWithMessage(Exception):
    """의도된 exception으로 traceback이 필요하지 않음"""
    def __init__(self, msg: str):
        Exception.__init__(self)
        self.__msg = msg

    def __str__(self):
        return self.__msg

def get_error_message(exc: Exception) -> str:
    """ExceptionWithMessage가 아니면 traceback 반환"""
    if isinstance(exc, ExceptionWithMessage):
        msg = str(exc)
    else:
        msg = f"예기치 않은 오류가 발생했습니다.\n\n{format_exc()}"
    return msg

def simplify_won(won: float, unit: str = "자동") -> str:
    """액수를 간편한 형태의 텍스트로 변환
    10,000 이상: 'OOO천원'
    10,000,000 이상: 'OOO백만원'

    Args:
        unit
            자동|억원|백만원|천원|원
    """
    if np.isnan(won):
        return ""
    prefix = "-" if won < 0 else ""
    won = abs(won)
    if (unit == "억원") or (unit == "자동" and won >= 100_000_000):
        return f"{prefix}{won/100_000_000:0,.1f}억원"
    if (unit == "백만원") or (unit == "자동" and won >= 1_000_000):
        return f"{prefix}{won/1_000_000:0,.1f}백만원"
    if (unit == "천원") or (unit == "자동" and won >= 1_000):
        return f"{prefix}{won/1_000:0,.1f}천원"
    return f"{prefix}{int(won):,}원"

COLORMAP = (
    "#F0B400",
    "#0A7771",
    "#263F66",
    "#6F0A73",
    "#9A9A9A"
)

def initialize_matplotlib():
    rc("font", family="Malgun Gothic")
    # font_manager.fontManager.addfont("./fonts/NotoSans-Regular.ttf")
    # plt.rcParams["font.family"] = [
    #     # "Malgun Gothic",
    #     font_manager.FontProperties(fname="./fonts/NotoSans-Regular.ttf").get_name(),
    # ]
    # plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams['axes.prop_cycle'] = plt.cycler(color=COLORMAP)

def pastel_gradient(hex_color: str, length: int, max_pastel: float = 0.7) -> list[str]:
    """기준 hex 색상에서 시작해서, 흰색 쪽으로 점점 파스텔톤으로 변하는 hex 색상 리스트를 반환"""
    if length <= 0:
        raise RuntimeError

    # '#' 있으면 제거
    hex_color = hex_color.lstrip("#")

    if len(hex_color) != 6:
        raise RuntimeError

    # 16진수 → 0~1 범위의 RGB
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    base_rgb = (r, g, b)

    def blend_with_white(rgb: tuple[float, float, float], t: float) -> tuple[float, float, float]:
        """
        rgb와 흰색(1,1,1)을 t 비율로 섞는다.
        t=0이면 원본 색, t=1이면 완전 흰색.
        """
        br, bg, bb = rgb
        wr, wg, wb = 1.0, 1.0, 1.0
        return (
            (1 - t) * br + t * wr,
            (1 - t) * bg + t * wg,
            (1 - t) * bb + t * wb,
        )

    def rgb_to_hex(rgb: tuple[float, float, float]) -> str:
        r, g, b = rgb
        return "#{:02X}{:02X}{:02X}".format(
            int(max(0, min(1, r)) * 255),
            int(max(0, min(1, g)) * 255),
            int(max(0, min(1, b)) * 255),
        )

    colors: list[str] = []

    if length == 1:
        # 하나만 만들면, 원본보다 살짝 파스텔톤으로 (max_pastel만큼 섞기)
        blended = blend_with_white(base_rgb, max_pastel)
        colors.append(rgb_to_hex(blended))
        return colors

    # 여러 개라면 t를 0 → max_pastel까지 균등하게 증가시키면서 생성
    for i in range(length):
        t = max_pastel * (i / (length - 1))  # 0, ..., max_pastel
        blended = blend_with_white(base_rgb, t)
        colors.append(rgb_to_hex(blended))

    return colors




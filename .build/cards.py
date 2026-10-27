import sys
sys.path.insert(0, "/workspace/.build")

from PIL import Image, ImageDraw
from ui import (
    Canvas, F, W, H, S, FONT_DISPLAY, FONT_BOLD, FONT_REG, FONT_MONO,
    WHITE, MUTED, DIM, GREEN, GREEN_DEEP, AMBER, RED,
    header, title, sub, footer, icon, phone,
)

OUT = "/workspace/闲鱼安卓定制"


def card1():
    c = Canvas.blank()
    header(c, "安卓定制 · 闲鱼接单")

    c.text((37, 96), "ANDROID  APP  DEVELOPMENT", F(FONT_MONO, 10), (120, 220, 175), tracking=2.2)
    title(c, "安卓 App", 116, 37)
    title(c, "定制开发", 166, 37, GREEN)
    sub(c, "小工具 / 单机本地 App / 功能定制 / 界面美化", 228, 37)

    c.text((37, 272), "¥", F(FONT_BOLD, 28), AMBER)
    c.text((55, 250), "199", F(FONT_DISPLAY, 74), AMBER)
    w199 = c.tw("199", F(FONT_DISPLAY, 74))
    c.text((55 + w199 + 6, 300), "起", F(FONT_BOLD, 14), AMBER)
    c.text((37, 344), "按功能报价 · 报价透明", F(FONT_REG, 13), MUTED)

    checks = ["可运行安装包，装机即用", "界面和功能按需求做", "验收期内 bug 免费修", "无广告、无弹窗，干净"]
    for i, t in enumerate(checks):
        y = 384 + i * 40
        c.circle(43, y + 8, 8, outline=GREEN + (120,), width=1.2)
        c.line(39.5, y + 8, 42, y + 11, GREEN + (255,), 1.6)
        c.line(42, y + 11, 47, y + 4, GREEN + (255,), 1.6)
        c.text((59, y), t, F(FONT_REG, 13), (206, 218, 238))

    c.rrect(37, 552, 236, 78, 12, fill=GREEN + (10,), outline=GREEN + (36,))
    c.text((54, 566), "不确定能不能做？", F(FONT_BOLD, 12.5), GREEN)
    c.text((54, 588), "先发需求，免费帮你判断", F(FONT_REG, 11), (185, 205, 218))
    c.text((54, 605), "不合适也会直说，不耽误你", F(FONT_REG, 11), (140, 160, 180))

    ph = phone().rotate(-7, expand=True, resample=Image.BICUBIC)
    c.img.alpha_composite(ph, (int(288 * S), int(272 * S)))
    c.draw = ImageDraw.Draw(c.img)

    footer(c, 1, "私信发需求 · 当天给报价 · 全程平台交易")
    c.save(f"{OUT}/01-封面.png")


def card2():
    c = Canvas.blank()
    header(c, "能做什么")
    title(c, "这些单子，", 92, 37)
    title(c, "我都能接", 146, 37, GREEN)
    sub(c, "只做单机本地 App，不用服务器、不用备案", 210, 37)

    items = [
        ("yen", "门店记账", "商品流水、日结月报"),
        ("box", "库存台账", "出入库登记、库存预警"),
        ("check", "巡检记录", "巡检项、周期提醒"),
        ("clock", "学员课时", "排课签到、课时扣减"),
        ("chart", "工时统计", "项目计时、成本汇总"),
        ("list", "保养计划", "到期提醒、历史记录"),
        ("grid", "健身训练", "组数重量、进度曲线"),
        ("page", "学习计划", "每日任务、完成率"),
    ]
    cw, ch, gx, gy = 221, 63, 23, 20
    x0, y0 = 37, 240
    for i, (kind, name, desc) in enumerate(items):
        col, row = i % 2, i // 2
        x = x0 + col * (cw + gx)
        y = y0 + row * (ch + gy)
        c.rrect(x, y, cw, ch, 12, fill=(255, 255, 255, 12), outline=(255, 255, 255, 26))
        icon(c, kind, x + 15, y + 14.5)
        c.text((x + 63, y + 13), name, F(FONT_BOLD, 15), WHITE)
        c.text((x + 63, y + 36), desc, F(FONT_REG, 10.5), (152, 166, 192))

    c.rrect(37, 562, 466, 54, 12, fill=GREEN + (10,), outline=GREEN + (34,))
    c.text((54, 574), "说明一下", F(FONT_BOLD, 12.5), GREEN)
    c.text((54, 595), "数据存手机本地，不联网也能用；要多人多设备同步的，我不接", F(FONT_REG, 10.5), (182, 202, 216))

    footer(c, 2, "功能越具体，报价越准 · 也可以先发参考 App")
    c.save(f"{OUT}/02-能做什么.png")


def card3():
    c = Canvas.blank()
    header(c, "价格")
    title(c, "按功能算价，", 92, 37)
    title(c, "报价透明", 146, 37, GREEN)
    sub(c, "把需求发我，当天给报价", 210, 37)

    tiers = [
        ("入门", "单页面小工具", "计算器 / 记账 / 清单 / 小工具", "周期 3–5 天", "199", False),
        ("标准", "多页面小应用", "多个页面 + 本地数据 + 界面美化", "周期 5–7 天", "399", True),
        ("进阶", "完整功能定制", "多模块 + 图表统计 + 复杂交互", "周期 7–15 天", "699", False),
    ]
    for i, (tier, name, desc, period, price, hot) in enumerate(tiers):
        y = 238 + i * 112
        fill = GREEN + (12,) if hot else (255, 255, 255, 10)
        edge = GREEN + (70,) if hot else (255, 255, 255, 28)
        c.rrect(37, y, 466, 104, 14, fill=fill, outline=edge)
        if hot:
            c.rrect(37, y, 3, 104, 1.5, fill=GREEN + (200,))

        c.text((57, y + 30), "¥", F(FONT_BOLD, 22), AMBER)
        c.text((72, y + 18), price, F(FONT_DISPLAY, 56), AMBER)
        wp = c.tw(price, F(FONT_DISPLAY, 56))
        c.text((72 + wp + 5, y + 60), "起", F(FONT_BOLD, 13), AMBER)

        lx = 202
        c.text((lx, y + 16), tier, F(FONT_BOLD, 12), GREEN if hot else (160, 174, 200))
        c.text((lx, y + 34), name, F(FONT_BOLD, 19), WHITE)
        c.text((lx, y + 62), desc, F(FONT_REG, 10.5), (156, 170, 196))
        c.text((lx, y + 82), period, F(FONT_REG, 10.5), (126, 140, 166))

        if hot:
            bw = c.tw("最多人问", F(FONT_BOLD, 9)) + 20
            c.rrect(503 - bw - 16, y + 16, bw, 21, 10.5, fill=GREEN + (36,))
            c.text((503 - bw - 16 + 10, y + 21), "最多人问", F(FONT_BOLD, 9), GREEN)

    c.text((37, 588), "为什么写「起」？", F(FONT_BOLD, 13), WHITE)
    c.text((37, 610), "功能数量差别很大，所以先给清单再报价；大改按新增功能另算", F(FONT_REG, 11), MUTED)

    footer(c, 3, "不接白菜价内卷，改动先说价")
    c.save(f"{OUT}/03-价格.png")


def card4():
    c = Canvas.blank()
    header(c, "合作流程")
    title(c, "四步走完，", 92, 37)
    title(c, "全程看得见", 146, 37, GREEN)
    sub(c, "每一步都有确认点，不用靠猜", 210, 37)

    steps = [
        ("发需求", "有文档发文档；没文档就说想法 + 一个参考 App，我帮你拆成功能点"),
        ("定方案", "功能清单 + 报价 + 工期，确认后开工"),
        ("看进度", "界面 → 功能 → 打包，中间随时可以问进度；要改提前说"),
        ("验收交付", "装机测试没问题 → 交安装包和使用说明 → 平台确认收货"),
    ]
    cy0, gap = 252, 80
    for i, (name, desc) in enumerate(steps):
        cy = cy0 + i * gap
        if i < len(steps) - 1:
            c.line(59, cy + 14, 59, cy + gap - 14, (255, 255, 255, 40), 1)
        c.circle(59, cy, 14, fill=(255, 255, 255, 8), outline=GREEN + (110,), width=1.2)
        num = "0%d" % (i + 1)
        wn = c.tw(num, F(FONT_MONO, 8.5))
        c.text((59 - wn / 2, cy - 5.5), num, F(FONT_MONO, 8.5), GREEN)
        c.text((90, cy - 13), name, F(FONT_BOLD, 17), WHITE)
        c.text((90, cy + 9), desc, F(FONT_REG, 10.5), (150, 165, 190))

    c.rrect(37, 574, 466, 42, 12, fill=GREEN + (10,), outline=GREEN + (34,))
    c.text((54, 587), "每天有进度，不做「消失的接单侠」", F(FONT_BOLD, 12), GREEN)

    footer(c, 4, "开工前把所有事说清楚，比事后扯皮省事")
    c.save(f"{OUT}/04-流程.png")


def card5():
    c = Canvas.blank()
    header(c, "接单范围")
    title(c, "什么接，", 92, 37)
    title(c, "什么不接", 146, 37, RED)
    sub(c, "边界说在前面，省得你白问我白忙", 210, 37)

    yes = ["单机本地 App", "小工具 / 效率工具", "企业内部小工具", "原型 Demo", "界面美化改造", "数据整理与统计"]
    no = ["外挂 / 破解 / 作弊", "抓取隐私数据", "赌博 / 擦边内容", "仿冒他人品牌", "刷量 / 刷单", "钓鱼 / 诈骗类"]

    def panel(y, label, items, col, h):
        c.rrect(37, y, 466, h, 14, fill=col + (10,), outline=col + (40,))
        c.circle(55, y + 23, 4, fill=col + (255,))
        c.text((67, y + 16), label, F(FONT_BOLD, 13), col)
        for i, t in enumerate(items):
            cx = 55 + (i // 3) * 218
            cy = y + 48 + (i % 3) * 30
            if col == GREEN:
                c.line(cx - 3, cy + 3, cx, cy + 6, col + (255,), 1.6)
                c.line(cx, cy + 6, cx + 5, cy - 2, col + (255,), 1.6)
            else:
                c.line(cx - 3, cy - 2, cx + 3, cy + 4, col + (255,), 1.6)
                c.line(cx - 3, cy + 4, cx + 3, cy - 2, col + (255,), 1.6)
            c.text((cx + 12, cy - 5), t, F(FONT_REG, 12), (206, 216, 236))

    panel(238, "可以接", yes, GREEN, 148)
    panel(398, "不接", no, RED, 148)

    c.text((37, 564), "只做单机本地 App：不用服务器、不用域名、不用备案。", F(FONT_REG, 11), (170, 185, 208))
    c.text((37, 584), "正规需求但拿不准？先把文档发来，我帮你看方向。", F(FONT_REG, 11), (170, 185, 208))

    c.rrect(37, 604, 466, 46, 14, fill=(255, 255, 255, 10), outline=(255, 255, 255, 26))
    c.text((54, 614), "私信格式", F(FONT_BOLD, 11), GREEN)
    c.text((54, 631), "需求一句话 + 参考 App + 截止时间（有文档直接发）", F(FONT_REG, 10.5), (160, 175, 200))

    footer(c, 5, "仅供参考，最终按平台规则和实际需求判断")
    c.save(f"{OUT}/05-接单范围.png")


def card6():
    c = Canvas.blank()
    header(c, "常见问题")
    title(c, "你可能想问的", 92, 37)
    title(c, "四个问题", 146, 37, GREEN)

    qa = [
        ("完全不懂安卓，装不上怎么办？",
         "安装包直接装手机就能用；权限怎么开、装不上怎么处理，安装说明里都写清楚。"),
        ("做出来和我说的不一样怎么办？",
         "开工前把功能和界面确认清楚；验收时按清单逐条对，对不上的我改。"),
        ("时间很赶，能加急吗？",
         "说明截止时间，排得上就接；排不上我会直接说，不接还拖着你最耽误事。"),
        ("中途改需求怎么算？",
         "开工前把清单定清楚；开工后加功能按新增项报价，报价确认再动手。"),
    ]
    for i, (q, a) in enumerate(qa):
        y = 218 + i * 82
        c.rrect(37, y - 2, 24, 24, 7, fill=GREEN + (40,))
        c.text((43, y + 2), "Q", F(FONT_BOLD, 12), GREEN)
        c.text((72, y - 2), q, F(FONT_BOLD, 15.5), WHITE)
        c.text((72, y + 24), a, F(FONT_REG, 10.5), (152, 166, 192))

    c.rrect(37, 566, 466, 54, 12, fill=GREEN + (10,), outline=GREEN + (34,))
    c.text((54, 578), "还是拿不准？", F(FONT_BOLD, 12.5), GREEN)
    c.text((54, 599), "把想法发我，先判断可行性，再谈价格", F(FONT_REG, 10.5), (182, 202, 216))

    footer(c, 6, "还有别的，直接私信，看到就回")
    c.save(f"{OUT}/06-常见问题.png")


if __name__ == "__main__":
    card1()
    card2()
    card3()
    card4()
    card5()
    card6()
    print("done")

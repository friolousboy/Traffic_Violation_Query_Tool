"""
违章查询脚本 — 从 sc.122.gov.cn 批量查询车辆违章信息。

读取 example.xlsx 中的车牌号，依次查询每个车牌的违章记录，
提取详细信息后保存到 output.xlsx。
"""

from __future__ import annotations

import sys
import time
from typing import Optional

class _Tee:
    """同时输出到控制台和日志文件。"""
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            f.write(obj)
            f.flush()
    def flush(self):
        for f in self.files:
            f.flush()

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait

# ============================================================
# 配置常量
# ============================================================
EXCEL_INPUT = "example.xlsx"
EXCEL_OUTPUT = "output.xlsx"
EXCEL_COLUMNS = ["车牌", "时间", "违法地点", "违法行为", "罚款金额", "记分值", "是否处理"]

HEADLESS = False            # True=后台静默（需先登录过且Cookie有效），False=显示浏览器
SAVE_INTERVAL = 20          # 每处理 N 个车牌保存一次
DEFAULT_TIMEOUT = 5        # 显式等待超时（秒）
LOGIN_TIMEOUT = 300         # 等待用户扫码登录的超时（秒）
PAGE_STABLE_WAIT = 5        # 翻页/查询后等表格稳定的最大秒数
DETAIL_DELAY = 0.5          # 详情面板打开后的稳定等待（秒）
MAX_ROW_RETRIES = 3         # 单条违章记录最大重试次数
MAX_PLATE_RETRIES = 3       # 单张车牌最大重新搜索次数

CAR_TYPES_ALL = ["小型汽车", "小型新能源汽车", "大型汽车"]

BASE_URL = "https://sc.122.gov.cn/views/memfyy/violation.html"

# ---- 元素定位器 ----
SELECTOR_CAR_TYPE = (By.ID, "hpzl")
SELECTOR_PLATE_INPUT = (By.ID, "hphm")
SELECTOR_SUBMIT = (By.CLASS_NAME, "btn-submit-veh")
SELECTOR_VIOLATION_ROWS = (By.CLASS_NAME, "surveil_view")
SELECTOR_CLOSE = (By.ID, "bind_close")

# 违章详情面板 — 字段名 → XPath
DETAIL_XPATHS: dict[str, str] = {
    "time":     '//*[@id="view"]/div[2]/div[1]/form/div[3]/span[2]',
    "location": '//*[@id="view"]/div[2]/div[1]/form/div[4]/span[2]',
    "behavior": '//*[@id="view"]/div[2]/div[1]/form/div[5]/span[2]',
    "fine":     '//*[@id="view"]/div[2]/div[1]/form/div[7]/span[2]',
    "points":   '//*[@id="view"]/div[2]/div[1]/form/div[8]/span[2]',
}



# ============================================================
# 爬虫类
# ============================================================
class ViolationScraper:
    """从 sc.122.gov.cn 抓取车辆违章信息。"""

    def __init__(self) -> None:
        """初始化 Edge 浏览器。"""
        options = EdgeOptions()
        if HEADLESS:
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
        self.driver: WebDriver = webdriver.Edge(options=options)

    # ---- 工具属性 ----
    @property
    def wait(self) -> WebDriverWait:
        """获取一个 WebDriverWait 实例（使用默认超时）。"""
        return WebDriverWait(self.driver, DEFAULT_TIMEOUT)

    # ---- 等待工具（带进度显示） ----
    @staticmethod
    def _tick(waited: int, total: int, label: str) -> None:
        """在终端同一行输出倒计时进度。"""
        print(f"  {label}... {waited}/{total}s", end="\r")

    def _poll(self, predicate, timeout: int, label: str = "等待") -> bool:
        """轮询等待条件成立，每秒打印进度。成功返回 True，超时返回 False。"""
        for i in range(timeout):
            try:
                if predicate():
                    print(f"  {label}... ✓ ({i}s)" + " " * 10)
                    return True
            except Exception:
                pass
            self._tick(i + 1, timeout, label)
            time.sleep(1)
        print(f"  {label}... 超时 ({timeout}s)" + " " * 10)
        return False

    # ========================================================
    # 登录处理
    # ========================================================
    def _is_on_query_page(self) -> bool:
        """检查当前页面是否是违章查询页（号牌输入框 + 查询按钮同时存在）。"""
        try:
            self.driver.find_element(*SELECTOR_PLATE_INPUT)
            self.driver.find_element(*SELECTOR_SUBMIT)
            return True
        except NoSuchElementException:
            return False

    def _navigate_to_violation_query(self) -> None:
        """从当前页面导航到违章查询页面。

        优先通过页面上的导航链接自动跳转；找不到时提示用户手动点击。
        """
        if self._is_on_query_page():
            return

        # 尝试点击常见导航链接
        nav_keywords = ["违法查询", "违章查询", "机动车违法", "违法处理"]
        for keyword in nav_keywords:
            try:
                link = self.driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
                link.click()
                time.sleep(2)
                if self._is_on_query_page():
                    print(f"已通过导航【{keyword}】进入查询页面")
                    return
            except NoSuchElementException:
                continue

        # 自动导航失败，请用户手动操作
        print("\n" + "=" * 50)
        print("请在浏览器中手动点击进入【机动车违法查询】页面")
        print("=" * 50)
        if self._poll(self._is_on_query_page, 120, "等待进入查询页面"):
            print("已进入查询页面 ✓")
            return
        raise TimeoutError("等待进入违章查询页面超时，请检查后重试")

    def _ensure_login(self) -> None:
        """打开网站等待用户扫码登录 → 选择公司 → 导航到违章查询页面。"""
        # 1. 打开登录页
        self.driver.get("https://gab.122.gov.cn/m/login?t=2")
        print("\n" + "=" * 50)
        print("请在浏览器中扫描二维码完成登录")
        print("=" * 50)

        def _logged_in() -> bool:
            url = self.driver.current_url
            if "login" in url.lower() or url.endswith("/"):
                return False
            try:
                return self.driver.find_element(By.TAG_NAME, "body").is_displayed()
            except Exception:
                return False

        if self._poll(_logged_in, LOGIN_TIMEOUT, "等待扫码登录"):
            print("登录成功 ✓")
        else:
            raise TimeoutError(f"等待登录超时（{LOGIN_TIMEOUT} 秒），请检查网络后重试")

        # 2. 等待用户点击公司（检测页面内容变化，兼容 SPA）
        print("\n" + "=" * 50)
        print("请在浏览器中点击【公司】进入")
        print("=" * 50)
        time.sleep(2)  # 等登录后的自动跳转完成
        old_body = self.driver.find_element(By.TAG_NAME, "body").text or ""
        if self._poll(
            lambda: (self.driver.find_element(By.TAG_NAME, "body").text or "") != old_body,
            300, "等待点击公司",
        ):
            print("已进入公司主页 ✓")
        else:
            print("未检测到页面跳转，继续...")

        # 3. 导航到违章查询页面
        self._navigate_to_violation_query()

    # ========================================================
    # 数据读写
    # ========================================================
    @staticmethod
    def read_plates(filepath: str = EXCEL_INPUT) -> list[str]:
        """从 Excel 读取车牌列表。"""
        df = pd.read_excel(filepath)
        return df["车牌"].tolist()

    @staticmethod
    def save_results(results: list[list], filepath: str = EXCEL_OUTPUT) -> None:
        """将累积结果保存到 Excel（去重后的唯一保存入口）。"""
        if not results:
            print("无数据可保存")
            return
        df = pd.DataFrame(results, columns=EXCEL_COLUMNS)
        print(df)
        df.to_excel(filepath, index=False)
        print(f"已保存 {len(results)} 条记录至 {filepath}")

    # ========================================================
    # 违章详情提取
    # ========================================================
    def _extract_detail_fields(self) -> dict[str, str]:
        """从违章详情弹窗中提取所有字段（单点真相）。"""
        fields: dict[str, str] = {}
        for name, xpath in DETAIL_XPATHS.items():
            try:
                el = self.wait.until(
                    EC.visibility_of_element_located((By.XPATH, xpath))
                )
                fields[name] = el.text
            except TimeoutException:
                fields[name] = ""
        return fields

    def _close_detail_panel(self) -> None:
        """安全关闭违章详情弹窗（点击 + JS 兜底）。"""
        try:
            btn = self.wait.until(EC.element_to_be_clickable(SELECTOR_CLOSE))
            btn.click()
        except (TimeoutException, ElementNotInteractableException):
            try:
                el = self.driver.find_element(*SELECTOR_CLOSE)
                self.driver.execute_script("arguments[0].click();", el)
            except Exception:
                pass  # 面板可能已关闭或不存���

    def _get_handled_status(self, row_index: int, total: int) -> str:
        """从违章列表表格读取"是否处理"列。"""
        if total == 1:
            xpath = '//*[@id="my-msg-list"]/tbody/tr/td[5]'
        else:
            xpath = f'//*[@id="my-msg-list"]/tbody/tr[{row_index + 1}]/td[5]'
        try:
            return self.wait.until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            ).text
        except TimeoutException:
            return ""

    def _dismiss_alert(self) -> None:
        """消除浏览器弹窗（如"请选择正确的号牌种类"）。"""
        try:
            alert = self.driver.switch_to.alert
            alert.accept()
        except Exception:
            pass

    def _wait_for_table_stable(self) -> None:
        """等待违章行表格稳定（行数不再变化），避免在渲染中途就处理数据。"""
        prev_count = -1
        deadline = time.time() + PAGE_STABLE_WAIT
        while time.time() < deadline:
            rows = self.driver.find_elements(*SELECTOR_VIOLATION_ROWS)
            cur_count = len(rows)
            if cur_count > 0 and cur_count == prev_count:
                return
            prev_count = cur_count
            elapsed = PAGE_STABLE_WAIT - max(0, deadline - time.time())
            self._tick(int(elapsed), PAGE_STABLE_WAIT, "表格稳定中")
            time.sleep(0.5)

    def _click_view_button(self, row_index: int) -> None:
        """点击指定行"操作"列中的「查看」按钮，打开违章详情面板。"""
        xpath = f'//*[@id="my-msg-list"]/tbody/tr[{row_index + 1}]/td[last()]/a'
        btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        # 滚到可见区域再点击
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        time.sleep(0.2)
        try:
            btn.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", btn)
        # 等待详情面板渲染完毕
        try:
            WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, DETAIL_XPATHS["location"]))
            )
        except TimeoutException:
            pass
        time.sleep(DETAIL_DELAY)

    # ========================================================
    # 单条违章记录处理
    # ========================================================
    def _process_violation_row(
        self, plate: str, row_index: int, total_rows: int
    ) -> Optional[list]:
        """点击一条违章的「查看」按钮，提取详情，返回结果行或 None。

        遇到元素过时（Stale）或提取为空时自动重试，最多 MAX_ROW_RETRIES 次。
        """
        handled = self._get_handled_status(row_index, total_rows)

        for attempt in range(MAX_ROW_RETRIES):
            try:
                self._click_view_button(row_index)

                fields = self._extract_detail_fields()

                if fields.get("location", "").strip():
                    self._close_detail_panel()
                    return [
                        plate,
                        fields["time"],
                        fields["location"],
                        fields["behavior"],
                        fields["fine"],
                        fields["points"],
                        handled,
                    ]

                # 关键字段为空，关闭面板后重试
                self._close_detail_panel()
                time.sleep(0.5)

            except StaleElementReferenceException:
                if attempt < MAX_ROW_RETRIES - 1:
                    continue
            except (ElementClickInterceptedException, ElementNotInteractableException):
                if attempt < MAX_ROW_RETRIES - 1:
                    continue
            except TimeoutException:
                if attempt < MAX_ROW_RETRIES - 1:
                    continue

        print(f"  违章行 {row_index + 1} 提取失败（已重试 {MAX_ROW_RETRIES} 次）")
        return None

    # ========================================================
    # 单车牌处理
    # ========================================================
    @staticmethod
    def _is_next_page_disabled(el) -> bool:
        """判断翻页元素是否处于禁用状态。"""
        classes = el.get_attribute("class") or ""
        return "disabled" in classes or "nolink" in classes

    def _go_to_next_page(self, page: int) -> bool:
        """尝试翻到下一页，成功返回 True，已到最后一页返回 False。"""
        def click_and_wait(el) -> bool:
            try:
                el.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", el)
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located(SELECTOR_VIOLATION_ROWS)
                )
            except TimeoutException:
                pass
            self._wait_for_table_stable()
            return True

        # 模式 1：「下一页」链接
        try:
            links = self.driver.find_elements(
                By.XPATH, '//a[contains(text(), "下一页")]'
            )
            for el in links:
                try:
                    if not self._is_next_page_disabled(el):
                        print(f"  → 翻至第{page + 1}页（下一页）")
                        return click_and_wait(el)
                except StaleElementReferenceException:
                    continue
            if links:
                return False  # 有"下一页"但被禁用 → 最后一页
        except Exception:
            pass

        # 模式 2：pagination 组件中的 » 或 > 符号
        try:
            pagination = self.driver.find_element(
                By.XPATH,
                '//*[contains(@class, "pagination") or contains(@class, "pager")]',
            )
            candidates = pagination.find_elements(
                By.XPATH, './/a[contains(text(), "»") or contains(text(), ">")]'
            )
            for el in candidates:
                try:
                    if not self._is_next_page_disabled(el):
                        print(f"  → 翻至第{page + 1}页（»）")
                        return click_and_wait(el)
                except StaleElementReferenceException:
                    continue
        except NoSuchElementException:
            pass

        # 模式 3：页码链接，点当前激活页的下一个兄弟
        try:
            active = self.driver.find_element(
                By.XPATH,
                '//*[contains(@class, "current") or contains(@class, "active")]',
            )
            parent = active.find_element(By.XPATH, "..")
            siblings = parent.find_elements(By.XPATH, "./a")
            for i, sib in enumerate(siblings):
                try:
                    if sib == active and i + 1 < len(siblings):
                        print(f"  → 翻至第{page + 1}页（页码）")
                        return click_and_wait(siblings[i + 1])
                except StaleElementReferenceException:
                    continue
        except NoSuchElementException:
            pass

        return False

    def _try_car_type(self, car_type: str, plate: str) -> list[list]:
        """选择号牌种类 → 点击查询 → 提取所有分页的违章记录。"""
        # ---- 等下拉框选项加载完毕 ----
        dropdown = self.wait.until(EC.element_to_be_clickable(SELECTOR_CAR_TYPE))
        self._poll(
            lambda: bool(dropdown.find_elements(By.TAG_NAME, "option"))
            and any(o.text for o in dropdown.find_elements(By.TAG_NAME, "option")),
            5, "等待下拉框就绪",
        )
        Select(dropdown).select_by_visible_text(car_type)

        # ---- 记下点击前的页面快照 ----
        old_body = self.driver.find_element(By.TAG_NAME, "body").text or ""

        # 点击查询按钮
        submit = self.wait.until(EC.element_to_be_clickable(SELECTOR_SUBMIT))
        submit.click()

        # ---- 等待页面刷新：body 内容必须变化（证明页面已响应查询） ----
        self._poll(
            lambda: (self.driver.find_element(By.TAG_NAME, "body").text or "") != old_body,
            3, "等待页面刷新",
        )

        # 等待新结果落定：违章行 或 「未查询到违法记录」
        def _got_rows() -> bool:
            return bool(self.driver.find_elements(*SELECTOR_VIOLATION_ROWS))

        def _got_empty() -> bool:
            return "未查询到违法记录" in (
                self.driver.find_element(By.TAG_NAME, "body").text or ""
            )

        if self._poll(_got_rows, 3, "等待违章数据加载"):
            outcome = "rows"
        elif _got_empty():
            outcome = "empty"
        else:
            self._poll(_got_empty, 3, "确认无违章")
            outcome = "empty" if _got_empty() else "timeout"

        if outcome != "rows":
            if outcome == "empty":
                print(f"  {plate}: {car_type} - 无违章")
            else:
                print(f"  {plate}: {car_type} - 查询超时，跳过")
            return []

        # 表格稳固后再处理，避免数据加载不全
        self._wait_for_table_stable()

        # ---- 逐页提取 ----
        all_results: list[list] = []
        page = 1

        while True:
            rows = self.driver.find_elements(*SELECTOR_VIOLATION_ROWS)
            if not rows:
                break

            print(f"  {plate}: {car_type} - 第{page}页, {len(rows)}条")
            for idx in range(len(rows)):
                row_data = self._process_violation_row(plate, idx, len(rows))
                if row_data:
                    all_results.append(row_data)
                    print(row_data)

            if not self._go_to_next_page(page):
                break
            page += 1

        return all_results

    def process_plate(self, plate: str) -> list[list]:
        """处理单张车牌，自动回退号牌种类。

        号牌种类尝试顺序：小型汽车 → 小型新能源汽车 → 大型汽车。
        弹窗时自动消除并尝试下一类型；关键元素缺失时重新搜索；
        点击持续失败时跳过该车牌。
        """
        plate_number = plate[1:]  # 去掉首位（省份简称）

        # 消除可能残留的弹窗
        self._dismiss_alert()

        # 等页面从上一次查询中恢复
        time.sleep(0.5)
        try:
            input_el = self.wait.until(EC.element_to_be_clickable(SELECTOR_PLATE_INPUT))
            input_el.clear()
            input_el.send_keys(plate_number)
        except UnexpectedAlertPresentException:
            self._dismiss_alert()
            # 弹窗消除后重试一次
            input_el = self.wait.until(EC.element_to_be_clickable(SELECTOR_PLATE_INPUT))
            input_el.clear()
            input_el.send_keys(plate_number)
        time.sleep(0.3)  # 等号牌种类下拉框响应输入

        # 根据车牌长度推断号牌种类顺序（新能源 8 位，蓝牌 7 位）
        if len(plate) == 8:
            types = ["小型新能源汽车", "小型汽车", "大型汽车"]
        else:
            types = ["小型汽车", "大型汽车"]

        for _ in range(MAX_PLATE_RETRIES):
            try:
                for car_type in types:
                    try:
                        return self._try_car_type(car_type, plate)
                    except UnexpectedAlertPresentException:
                        self._dismiss_alert()
                        print(f"  {plate}: {car_type} 触发弹窗，尝试下一类型")
                        continue
                print(f"  {plate}: 所有类型均弹窗，重新搜索")
            except NoSuchElementException:
                print(f"  {plate}: 页面元素缺失，重新搜索")
            except ElementClickInterceptedException:
                print(f"  {plate}: 点击持续被拦截，跳过")
                return []

        print(f"  {plate}: 超过最大重试次数（{MAX_PLATE_RETRIES}），跳过")
        return []

    # ========================================================
    # 生命周期
    # ========================================================
    def run(self) -> None:
        """主入口：读取车牌列表 → 逐条查询 → 定期保存 → 最终保存。"""
        plates = self.read_plates()
        all_results: list[list] = []

        self._ensure_login()

        for idx, plate in enumerate(plates):
            print(f"\n{'='*40}")
            print(f"开始搜索 [{idx + 1}/{len(plates)}] {plate}")
            print(f"{'='*40}")

            plate_results = self.process_plate(plate)
            all_results.extend(plate_results)

            # 每 N 条增量保存
            if (idx + 1) % SAVE_INTERVAL == 0:
                self.save_results(all_results)

        # 最终保存
        print(f"\n{'='*40}")
        print("查询完成，最终保存...")
        print(f"{'='*40}")
        self.save_results(all_results)

    def quit(self) -> None:
        """安全关闭浏览器。"""
        try:
            self.driver.quit()
        except Exception:
            pass


# ============================================================
# 入口
# ============================================================
if __name__ == "__main__":
    # 同时输出到控制台和日志文件
    log = open("log.txt", "w", encoding="utf-8")
    sys.stdout = _Tee(sys.stdout, log)
    sys.stderr = _Tee(sys.stderr, log)

    scraper = ViolationScraper()
    try:
        scraper.run()
    except Exception as e:
        print(f"\n[FATAL] {e}")
        import traceback
        traceback.print_exc()
    finally:
        scraper.quit()
        print("脚本结束。")
        log.close()

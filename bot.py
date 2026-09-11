import json
import os
import asyncio
from datetime import datetime, timedelta, time as datetime_time
import pandas as pd
import telegram 
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, JobQueue
from zoneinfo import ZoneInfo
import time as time_module
import shutil

# ================== 配置区 ==================
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise ValueError("❌ 请在 Railway 设置环境变量 BOT_TOKEN")

print(f"✅ Bot Token 已成功加载 | 长度: {len(TOKEN)}")

# ================== Railway Volume 配置 ==================
DATA_PATH = os.getenv("DATA_PATH", "/data")
os.makedirs(DATA_PATH, exist_ok=True)

DATA_FILE = os.path.join(DATA_PATH, "group_attendance.json")
EXCEL_FOLDER = os.path.join(DATA_PATH, "excel_files")
os.makedirs(EXCEL_FOLDER, exist_ok=True)

# ================== 北京时间 ==================
TZ = ZoneInfo("Asia/Shanghai")

def beijing_now():
    return datetime.now(TZ)

def beijing_date_str(dt=None):
    if dt is None:
        dt = beijing_now()
    return dt.strftime("%Y-%m-%d")



# ================== 时间有效性检查 ==================
def is_valid_checkin_time(shift: str, now: datetime = None) -> tuple[bool, str]:
    if shift not in {"1", "2", "3", "4"}:
        return True, ""
    
    if now is None:
        now = beijing_now()
    
    current_time = now.time()
    
    if shift == "1":      
        if current_time < datetime_time(11, 30):
            return False, "⚠️ 第一班上班需在 **11:30之后** 打卡"
    elif shift == "2":    
        if current_time >= datetime_time(18, 30):
            return False, "⚠️ 第一班下班需在 **18:30之前** 打卡"
    elif shift == "3":    
        if current_time < datetime_time(18, 30):
            return False, "⚠️ 第二班上班需在 **18:30之后** 打卡"
    elif shift == "4":    
        if current_time >= datetime_time(3, 30):
            return False, "⚠️ 第二班下班需在 **03:30之前** 打卡（00:00-03:30）"
    
    return True, ""


def is_valid_rest_time(shift: str) -> tuple[bool, str]:
    if shift not in {"5", "7"}:
        return True, ""
    
    now = beijing_now()
    current_time = now.time()
    
    if (datetime_time(12, 00) <= current_time < datetime_time(18, 0)) or \
       (datetime_time(19, 00) <= current_time) or \
       (current_time < datetime_time(3, 0)):
        return True, ""
    
    return False, "⚠️ 休息/暂离（5或7）只能在以下工作时段打卡：\n• 第一班 12:00-18:00\n• 第二班 19:00-03:00"


def calculate_rest_duration(start_time_str: str, end_time_str: str) -> int:
    try:
        fmt = "%H:%M:%S"
        start = datetime.strptime(start_time_str, fmt)
        end = datetime.strptime(end_time_str, fmt)
        
        if end < start:
            end += timedelta(days=1)
        
        delta = end - start
        return int(delta.total_seconds() / 60)
    except Exception as e:
        print(f"⚠️ 计算休息时长失败: {e}")
        return 0

# ================== 【状态判断工具函数】==================
def get_open_status(records: list) -> tuple[bool, bool]:
    """严格判断是否存在未结束的休息和暂离 - 修复版"""
    open_rest = False
    open_work_rest = False
    
    if not records:
        return False, False
    
    # 查找最近的未结束休息 (5)
    latest_rest_start_index = -1
    latest_rest_end_index = -1
    for i, r in enumerate(reversed(records)):
        idx = len(records) - 1 - i # 原始索引
        act = str(r.get("action", ""))
        if act == "5":
            if latest_rest_start_index == -1: # 找到最近的5
                latest_rest_start_index = idx
            # 如果这个5没有rest_minutes，说明是未结束的
            if "rest_minutes" not in r:
                open_rest = True
                break # 找到未结束的5，可以确定状态并退出
        elif act == "6":
            if latest_rest_end_index == -1: # 找到最近的6
                latest_rest_end_index = idx
            # 如果最近的6比最近的5晚，说明5已结束
            if latest_rest_start_index != -1 and latest_rest_end_index > latest_rest_start_index:
                open_rest = False
                break # 找到已结束的5，可以确定状态并退出
    
    # 查找最近的未结束暂离 (7)
    latest_work_rest_start_index = -1
    latest_work_rest_end_index = -1
    for i, r in enumerate(reversed(records)):
        idx = len(records) - 1 - i # 原始索引
        act = str(r.get("action", ""))
        if act == "7":
            if latest_work_rest_start_index == -1: # 找到最近的7
                latest_work_rest_start_index = idx
            # 如果这个7没有rest_minutes，说明是未结束的
            if "rest_minutes" not in r:
                open_work_rest = True
                break # 找到未结束的7，可以确定状态并退出
        elif act == "8":
            if latest_work_rest_end_index == -1: # 找到最近的8
                latest_work_rest_end_index = idx
            # 如果最近的8比最近的7晚，说明7已结束
            if latest_work_rest_start_index != -1 and latest_work_rest_end_index > latest_work_rest_start_index:
                open_work_rest = False
                break # 找到已结束的7，可以确定状态并退出
                
    return open_rest, open_work_rest


# ================== 状态判断工具函数 ==================
def is_currently_on_duty(records: list) -> bool:
    """判断当前是否在岗（可保留，暂时未被调用）"""
    if not records:
        return False

    shift1_active = False
    shift2_active = False

    for r in records:
        act = r.get("action")
        if act == "1":
            shift1_active = True
        elif act == "2":
            shift1_active = False
        elif act == "3":
            shift2_active = True
        elif act == "4":
            shift2_active = False

    return shift1_active or shift2_active


def has_started_work_today(records: list) -> bool:
    return any(r.get("action") in {"1", "3"} for r in records)


def get_late_minutes(expected: str, shift: str = None, now: datetime = None) -> tuple[int, str]:
    if not expected or shift not in {"1", "3"}:
        return 0, ""
    
    if now is None:
        now = beijing_now()
    
    try:
        exp_hm = datetime.strptime(expected, "%H:%M").time()
        
        expected_dt = now.replace(hour=exp_hm.hour, minute=exp_hm.minute, 
                                second=0, microsecond=0)
        
        if shift == "3":
            if now.hour < 1:  
                expected_dt -= timedelta(days=1)
        
        delta = now - expected_dt
        late_seconds = max(0, int(delta.total_seconds()))
        
        if late_seconds == 0:
            return 0, ""
        elif late_seconds < 60:
            return late_seconds, f"（迟到{late_seconds}秒）"
        else:
            late_min = late_seconds // 60
            return late_min, f"（迟到{late_min}分钟）"
            
    except Exception as e:
        print(f"迟到计算异常: {e}")
        return 0, ""

# ================== 日期逻辑 ==================
def get_attendance_date(now=None):
    """考勤日期：04:00 为分界点，所有00:00-03:59的记录都算前一天"""
    if now is None:
        now = beijing_now()
    if now.hour < 4:   # 00:00-03:59 算前一天
        return beijing_date_str(now - timedelta(days=1))
    return beijing_date_str(now)

def get_record_date(shift: str, now=None) -> str:
    """所有打卡记录的日期都以 get_attendance_date 为准"""
    if now is None:
        now = beijing_now()
    # 统一使用 get_attendance_date 的逻辑，不再区分 shift
    return get_attendance_date(now)

# ================== 日报日期函数 ==================
def get_report_date_for_daily(now=None) -> str:
    """自动日报专用 - 确保取前一天完整考勤日期 (基于04:00分界点)"""
    if now is None:
        now = beijing_now()
    # 在 04:00 之前触发日报，应该生成前一天的报表
    # 如果当前时间是 01:30，那么 get_attendance_date(now) 会返回前一天
    # 所以直接返回 get_attendance_date(now) 即可，它会根据 02:00 分界点自动处理
    return get_attendance_date(now)

# ================== DataManager ==================
class DataManager:
    def __init__(self):
        self._data: dict = {}
        self._last_mtime = 0
        self._dirty = False
        self._global_lock = asyncio.Lock()
        self._chat_locks: dict[str, asyncio.Lock] = {}
        self._save_task = None

    def _get_chat_lock(self, chat_id: str):
        if chat_id not in self._chat_locks:
            self._chat_locks[chat_id] = asyncio.Lock()
        return self._chat_locks[chat_id]

    def _file_mtime(self) -> float:
        try:
            return os.path.getmtime(DATA_FILE) if os.path.exists(DATA_FILE) else 0
        except:
            return 0

    def load(self, force: bool = False) -> dict:
        current_mtime = self._file_mtime()
        if force or current_mtime > self._last_mtime or not self._data:
            if os.path.exists(DATA_FILE):
                try:
                    with open(DATA_FILE, "r", encoding="utf-8") as f:
                        self._data = json.load(f)
                    print(f"📥 数据加载完成 | 群组数: {len(self._data)}")
                except Exception as e:
                    print(f"❌ 加载失败: {e}")
                    self._data = {}
            else:
                self._data = {}
            self._last_mtime = current_mtime
            self._dirty = False
        return self._data

    async def aload(self, force: bool = False):
        return await asyncio.to_thread(self.load, force)

    def _sync_save(self, data):
        try:
            temp_file = DATA_FILE + ".tmp"
            backup_file = DATA_FILE + ".bak"
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if os.path.exists(DATA_FILE):
                shutil.copy2(DATA_FILE, backup_file)
            os.replace(temp_file, DATA_FILE)
            return True
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return False

    async def save(self, immediate: bool = False):
        async with self._global_lock:
            if not self._dirty and not immediate:
                return
            success = await asyncio.to_thread(self._sync_save, self._data.copy())
            if success:
                self._last_mtime = self._file_mtime()
                self._dirty = False
                print(f"💾 数据已保存 | 群组: {len(self._data)}")

    async def _delayed_save(self):
        await asyncio.sleep(3)
        await self.save()

    async def get_chat_data(self, chat_id: str):
        chat_id_str = str(chat_id)
        if not chat_id_str.startswith("-"):
            return {"registered": {}, "users": {}, "admins": [], "activated": False}
        
        async with self._get_chat_lock(chat_id_str):
            await self.aload()  # 轻量读取
            return self._data.setdefault(chat_id_str, {
                "registered": {}, "users": {}, "admins": [], "activated": False
            })

    async def update_chat_data(self, chat_id: str, chat_data: dict):
        chat_id_str = str(chat_id)
        if not chat_id_str.startswith("-"):
            return
        async with self._get_chat_lock(chat_id_str):
            self._data[chat_id_str] = chat_data
            self._dirty = True
            if not self._save_task or self._save_task.done():
                self._save_task = asyncio.create_task(self._delayed_save())

    async def force_save(self):
        await self.save(immediate=True)

    async def cleanup_old_data(self, context: ContextTypes.DEFAULT_TYPE = None):
        async with self._global_lock:
            await self.aload(force=True)
            today = get_attendance_date(beijing_now())
            cutoff = get_attendance_date(beijing_now() - timedelta(days=35))
            cleaned = 0
            for chat_id_str in list(self._data.keys()):
                if not chat_id_str.startswith('-'):
                    continue
                for user_id in list(self._data[chat_id_str].get("users", {}).keys()):
                    records = self._data[chat_id_str]["users"][user_id].get("records", {})
                    for d in list(records.keys()):
                        if d < cutoff and d != today:
                            del records[d]
                            cleaned += 1
            if cleaned > 0:
                self._dirty = True
                await self.force_save()
                print(f"🧹 清理了 {cleaned} 条旧记录")

# ================== ACTIONS ==================
ACTIONS = {
    "1": {"name": "第一班上班", "time": "12:30", "is_work": True,  "type": "work"},
    "2": {"name": "第一班下班", "time": "17:30", "is_work": False, "type": "work"},
    "3": {"name": "第二班上班", "time": "19:30", "is_work": True,  "type": "work"},
    "4": {"name": "下班打卡完成，请将设备摆放整齐并开启飞行模式", "time": "02:30", "is_work": False, "type": "work"},
    "5": {"name": "开始休息",       "time": None, "is_work": False, "type": "rest_start"},
    "6": {"name": "结束休息",       "time": None, "is_work": False, "type": "rest_end"},
    "7": {"name": "工作原因暂离座位", "time": None, "is_work": False, "type": "work_rest_start"},
    "8": {"name": "作业结束回到座位", "time": None, "is_work": False, "type": "work_rest_end"},
}


# ================== 新增：群组激活检查 ==================
def is_group_activated(chat_id: str, chat_data: dict = None) -> bool:
    """检查群组是否已激活"""
    if chat_data:
        return chat_data.get("activated", False)
    
    # 从 data_manager 获取（安全方式）
    try:
        if hasattr(data_manager, '_data'):
            return data_manager._data.get(str(chat_id), {}).get("activated", False)
    except:
        pass
    return False


# ================== 休息超时提醒（群组提醒） ==================
async def check_rest_timeout(context: ContextTypes.DEFAULT_TYPE):
    job_data = context.job.data
    chat_id = job_data["chat_id"]
    user_id = job_data["user_id"]
    start_time = job_data.get("start_time")
    
    try:
        chat_data = await data_manager.get_chat_data(str(chat_id))
        user_name = chat_data.get("registered", {}).get(user_id, "未知用户")
        
        reminder_text = f"⚠️ **休息超时提醒**\n\n" \
                       f"👤 {user_name}\n" \
                       f"🕒 您已在 **{start_time}** 开始休息，已超过 **60分钟** 仍未结束休息。\n\n" \
                       f"请尽快回复 **6** 结束休息！"

        await context.bot.send_message(
            chat_id=chat_id, 
            text=reminder_text, 
            parse_mode="Markdown"
        )
            
    except Exception as e:
        print(f"休息提醒异常: {e}")


# ================== 报表生成 ==================
def build_daily_report_rows(chat_data: dict, report_date: str):
    """日报 - 单日报表（已修复多记录覆盖问题）"""
    
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})
    
    rows = []
    
    for user_id, user_name in registered.items():
        user_info = users.get(user_id, {"name": user_name, "records": {}})
        records_dict = user_info.get("records", {})
        records: list = records_dict.get(report_date, [])
        
        # ================== 修复重点 ==================
        # 按 action 收集最后一条有效记录（更可靠）
        shifts = {}
        for r in records:
            act = r.get("action")
            if act in {"1", "2", "3", "4"}:
                shifts[act] = r   # 后面出现的会覆盖前面 → 保留最后一次
        
        # ================== 休息统计 ==================
        total_rest = 0
        rest_count = 0
        total_work_rest = 0
        work_rest_count = 0

        for r in records:
            minutes = r.get("rest_minutes")
            if minutes is not None:
                action = r.get("action")
                if action == "6":
                    total_rest += max(0, minutes or 0)
                    rest_count += 1
                elif action == "8":
                    total_work_rest += max(0, minutes or 0)
                    work_rest_count += 1

        late1 = shifts.get("1", {}).get("late_display", "")
        late2 = shifts.get("3", {}).get("late_display", "")

        missing = set('1234') - set(shifts.keys())
        status = "正常" if not missing else f"缺卡: {','.join(sorted(missing))}"

        rows.append({
            "姓名": user_name,
            "日期": report_date,
            "第一班上班": shifts.get("1", {}).get("time", "缺卡"),
            "第一班下班": shifts.get("2", {}).get("time", "缺卡"),
            "第二班上班": shifts.get("3", {}).get("time", "缺卡"),
            "第二班下班": shifts.get("4", {}).get("time", "缺卡"),
            "第一班迟到": late1,
            "第二班迟到": late2,
            "休息次数": rest_count,
            "总休息分钟": total_rest,
            "工作原因休息次数": work_rest_count,
            "工作原因总休息分钟": total_work_rest,
            "状态": status,
        })
    
    rows.sort(key=lambda x: x["姓名"])

    return rows

def build_month_report_rows(chat_data: dict, month: str):
    """月报表 - 按天展开（修复版）"""

    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})
    rows = []
    
    for user_id, user_name in registered.items():
        user_records = users.get(user_id, {}).get("records", {})
        
        for date, records in user_records.items():
            if not date.startswith(month):
                continue
            
            # ================== 核心修复 ==================
            shifts = {}
            for r in records:
                act = str(r.get("action"))
                if act in {"1", "2", "3", "4"}:
                    shifts[act] = r   # 保留最后一条
            
            # ================== 休息统计 ==================
            total_rest = 0
            rest_count = 0
            total_work_rest = 0
            work_rest_count = 0

            for r in records:
                minutes = r.get("rest_minutes")
                if minutes is not None:
                    action = str(r.get("action"))
                    if action == "6":
                        total_rest += max(0, minutes or 0)
                        rest_count += 1
                    elif action == "8":
                        total_work_rest += max(0, minutes or 0)
                        work_rest_count += 1

            late1 = shifts.get("1", {}).get("late_display", "")
            late2 = shifts.get("3", {}).get("late_display", "")

            missing = set('1234') - set(shifts.keys())
            status = "正常" if not missing else f"缺卡: {','.join(sorted(missing))}"

            rows.append({
                "姓名": user_name,
                "日期": date,
                "第一班上班": shifts.get("1", {}).get("time", "缺卡"),
                "第一班下班": shifts.get("2", {}).get("time", "缺卡"),
                "第二班上班": shifts.get("3", {}).get("time", "缺卡"),
                "第二班下班": shifts.get("4", {}).get("time", "缺卡"),
                "第一班迟到": late1,
                "第二班迟到": late2,
                "休息次数": rest_count,
                "总休息分钟": total_rest,
                "工作原因休息次数": work_rest_count,
                "工作原因总休息分钟": total_work_rest,
                "状态": status,
            })
    
    rows.sort(key=lambda x: (x["姓名"], x["日期"]))

    return rows

def cleanup_old_excels():
    try:
        now = beijing_now()
        for f in os.listdir(EXCEL_FOLDER):
            if f.endswith(".xlsx"):
                path = os.path.join(EXCEL_FOLDER, f)
                file_mtime = datetime.fromtimestamp(os.path.getmtime(path), tz=TZ)
                if (now - file_mtime).days >= 3:
                    os.remove(path)
                    print(f"🗑️ 已清理过期Excel: {f}")
    except Exception as e:
        print(f"清理Excel失败: {e}")


# ================== 核心打卡函数（已增加单次限制）==================
async def daka(update: Update, context: ContextTypes.DEFAULT_TYPE, shift: str):
    chat_id_str = str(update.effective_chat.id)
    user = update.effective_user
    user_id = str(user.id)

    # ================== 激活检查 ==================
    if update.effective_chat.type != "private":
        chat_data_temp = await data_manager.get_chat_data(chat_id_str)
        if not chat_data_temp.get("activated", False):
            await update.message.reply_text(
                "⚠️ **本群尚未激活**\n\n此机器人需要密码才能激活使用。\n请联系机器人管理员获取激活密码。",
                parse_mode="Markdown"
            )
            return

    await auto_register(update, context)

    now = beijing_now()
    date_str = get_record_date(shift, now)

    # ================== 私聊禁用打卡 ==================
    if update.effective_chat.type == "private":
        await update.message.reply_text(
            "⚠️ **私聊中无法使用打卡功能**\n\n请在群聊中发送 1-8 进行打卡。",
            parse_mode="Markdown"
        )
        return

    # ================== 有效性检查 ==================
    valid, msg = is_valid_checkin_time(shift, now)
    if not valid:
        await update.message.reply_text(msg)
        return

    if shift in ["5", "7"]:
        valid_rest, rest_msg = is_valid_rest_time(shift)
        if not valid_rest:
            await update.message.reply_text(rest_msg)
            return

    # 加载数据
    chat_data = await data_manager.get_chat_data(chat_id_str)
    user_data = chat_data["users"].setdefault(user_id, {"name": user.full_name, "records": {}})
    records: list = user_data["records"].setdefault(date_str, [])

    action_info = ACTIONS.get(shift, {"name": f"操作{shift}", "type": "unknown"})

    # ================== 【新增】1,2,3,4 单次打卡限制 ==================
    if shift in {"1", "2", "3", "4"}:
        for r in records:
            if r.get("action") == shift:
                await update.message.reply_text(
                    f"⚠️ **{user.full_name}**\n\n"
                    f"今日 **{action_info['name']}** 已打卡，无需重复打卡！\n"
                    f"时间：{r.get('time', '未知')}",
                    parse_mode="Markdown"
                )
                return

    # ================== 状态判断 ==================
    open_rest, open_work_rest = get_open_status(records)

    # ================== 严格业务规则 ==================
    if shift == "5":
        if open_rest or open_work_rest:
            await update.message.reply_text("⚠️ 当前有未结束的休息/暂离，请先结束后再开始新的")
            return

    if shift == "7":
        if open_rest or open_work_rest:
            await update.message.reply_text("⚠️ 当前有未结束的休息/暂离，请先结束后再暂离")
            return

    if shift == "6":
        if not open_rest:
            await update.message.reply_text("⚠️ 请先输入5开始休息")
            return

    if shift == "8":
        if not open_work_rest:
            await update.message.reply_text("⚠️ 请先输入7开始暂离")
            return

    # 下班必须结束休息
    if shift in ["2", "4"]:
        if open_rest or open_work_rest:
            await update.message.reply_text("⚠️ 下班前必须先输入6/8 结束休息/暂离")
            return

    # ================== 执行打卡 ==================
    now_time_str = now.strftime("%H:%M:%S")
    late_seconds, late_txt = get_late_minutes(action_info.get("time"), shift, now)
    final_display = action_info["name"]

    if shift in ["6", "8"]:
        target = "5" if shift == "6" else "7"
        matched = False
        for r in reversed(records):
            if r.get("action") == target and "rest_minutes" not in r:
                rest_min = calculate_rest_duration(r["time"], now_time_str)
                final_display = f"{action_info['name']}（{rest_min}分钟）"
                
                records.append({
                    "time": now_time_str,
                    "action": shift,
                    "display": final_display,
                    "rest_minutes": rest_min,
                    "type": action_info.get("type")
                })
                r["rest_minutes"] = rest_min
                matched = True

                if shift == "6" and r.get("rest_job_name"):
                    for job in list(context.job_queue.get_jobs_by_name(r["rest_job_name"])):
                        job.schedule_removal()
                break
        if not matched:
            await update.message.reply_text(f"⚠️ 未找到对应开始记录")
            return
    else:
        record_entry = {
            "time": now_time_str,
            "action": shift,
            "display": action_info["name"],
            "type": action_info.get("type")
        }
        if late_seconds > 0 and shift in ["1", "3"]:
            record_entry["late_seconds"] = late_seconds
            record_entry["late_display"] = late_txt
            record_entry["display"] = f"{action_info['name']}{late_txt}"
            final_display = record_entry["display"]

        records.append(record_entry)

        if shift == "5":
            job_name = f"rest_timeout_{chat_id_str}_{user_id}_{int(time_module.time())}"
            record_entry["rest_job_name"] = job_name
            context.job_queue.run_once(check_rest_timeout, 3600, data={
                "chat_id": int(chat_id_str), "user_id": user_id, "start_time": now_time_str, "job_name": job_name
            }, name=job_name)

    await data_manager.update_chat_data(chat_id_str, chat_data)

    emoji = "⚠️" if late_seconds > 0 else "✅"
    await update.message.reply_text(
        f"{emoji} **{user.full_name}** {final_display}\n日期：{date_str}\n时间：{now_time_str}",
        parse_mode="Markdown"
    )

# ================== 消息处理 ==================
async def text_daka(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """文本打卡处理 - 私聊禁用1-8打卡"""
    # ================== 私聊禁用打卡 ==================
    if update.effective_chat.type == "private":
        await update.message.reply_text(
        "飞机的代号确定下来了就不要再改了，否则会打卡记录失败\n\n"
        "第一班上班打1，下班打2。第二班上班打3，下班打4，离开工位休息打5，回来打6（离开工位没打卡1次50，不论任何原因）\n\n"
        "上下班打卡的，迟到早退相同，10分钟内扣50，1小时内扣100，1小时外按旷工扣200。上班根据机器人的打卡时间，超过1秒也算迟到。漏打卡每次100\n"
        "⚠️严禁互相打卡与飞机定时发送。互相打卡两个人各扣300，定时发送扣600⚠️\n"
        "下班没打卡的不管是加班聊客户或者其他原因没打卡的一律算漏打卡。（下班打卡有效时间1小时）\n\n"
        "私聊机器人发送 /myrecord 可查询个人打卡记录\n"
    )
        return

    # 仅群聊才执行打卡逻辑
    text = update.message.text.strip().lower()
    mapping = {
        "1":"1","上班":"1","上午":"1",
        "2":"2","下班":"2","下班1":"2","下1":"2",
        "3":"3","下午上班":"3","上班2":"3",
        "4":"4","下班2":"4","下2":"4",
        "5":"5","休息":"5","开始休息":"5",
        "6":"6","结束休息":"6","回岗":"6",
        "7":"7","暂离":"7","离开":"7","工作原因休息":"7",
        "8":"8","回到座位":"8","回座位":"8",
    }
    if text in mapping:
        await daka(update, context, mapping[text])


async def auto_register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """仅在群聊中自动注册用户，私聊不创建数据"""
    if update.effective_chat.type == "private":
        return  # 私聊不自动注册，也不创建群组数据

    chat_id_str = str(update.effective_chat.id)
    user_id = str(update.effective_user.id)
    name = update.effective_user.full_name

    chat_data = await data_manager.get_chat_data(chat_id_str)
    if user_id not in chat_data["registered"]:
        chat_data["registered"][user_id] = name
        chat_data["users"].setdefault(user_id, {"name": name, "records": {}})
        await data_manager.update_chat_data(chat_id_str, chat_data)
        # await update.message.reply_text(f"✅ **{name}** 自动注册成功！", parse_mode="Markdown")


# ================== 管理员权限判断 ==================
async def is_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    if update.effective_chat.type == "private":
        return True

    user_id = str(update.effective_user.id)
    chat_id_str = str(update.effective_chat.id)

    try:
        member = await update.effective_chat.get_member(update.effective_user.id)
        if member.status in ["administrator", "creator"]:
            return True
    except:
        pass

    try:
        chat_data = await data_manager.get_chat_data(chat_id_str)
        if user_id in chat_data.get("admins", []):
            return True
    except:
        pass

    return False


async def get_group_owner(context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    try:
        admins = await context.bot.get_chat_administrators(chat_id)
        for a in admins:
            if a.status == "creator":
                return str(a.user.id)
    except:
        pass
    return None


# ================== 激活命令 ==================
async def activate_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """使用密码激活群组"""
    if update.effective_chat.type == "private":
        await update.message.reply_text("✅ 私聊无需激活，直接使用即可。")
        return

    chat_id_str = str(update.effective_chat.id)
    
    if not context.args:
        await update.message.reply_text(
            "❌ 格式错误\n\n"
            "正确用法：`/secretactivate 你的密码`",
            parse_mode="Markdown"
        )
        return

    secret_code = context.args[0].strip()
    CORRECT_SECRET = "acai888"   # ←←← 这里改成你想要的专属密码

    if secret_code != CORRECT_SECRET:
        await update.message.reply_text("❌ 密码错误，无权激活！")
        return

    chat_data = await data_manager.get_chat_data(chat_id_str)
    
    if chat_data.get("activated"):
        await update.message.reply_text("✅ 本群已激活，无需重复操作。")
        return

    chat_data["activated"] = True
    await data_manager.update_chat_data(chat_id_str, chat_data)
    await data_manager.force_save()
    
    await update.message.reply_text(
        "🎉 **本群已成功激活**！\n\n"
        "所有成员现在可以正常使用打卡功能（发送1-8）。\n"
        "激活永久有效。",
        parse_mode="Markdown"
    )


# ================== 管理员命令 ==================
async def add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return

    chat_id_str = str(update.effective_chat.id)
    operator_id = str(update.effective_user.id)
    owner_id = await get_group_owner(context, int(chat_id_str))

    if owner_id and owner_id != operator_id:
        await update.message.reply_text("⚠️ 仅群主可添加/删除管理员")
        return

    if not context.args:
        await update.message.reply_text("用法: `/addadmin <用户ID>`\n例如: `/addadmin 123456789`", parse_mode="Markdown")
        return

    target = context.args[0].strip()
    if not target.isdigit():
        await update.message.reply_text("❌ 用户ID必须为纯数字（如：123456789）")
        return

    chat_data = await data_manager.get_chat_data(chat_id_str)
    admins = chat_data.setdefault("admins", [])

    if target in admins:
        await update.message.reply_text("✅ 该用户已是管理员")
        return

    admins.append(target)
    await data_manager.update_chat_data(chat_id_str, chat_data)
    await data_manager.force_save()

    await update.message.reply_text(f"✅ 已成功添加管理员\n👤 ID: `{target}`", parse_mode="Markdown")


async def del_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return

    chat_id_str = str(update.effective_chat.id)
    operator_id = str(update.effective_user.id)
    owner_id = await get_group_owner(context, int(chat_id_str))

    if owner_id and owner_id != operator_id:
        await update.message.reply_text("⚠️ 仅群主可添加/删除管理员")
        return

    if not context.args:
        await update.message.reply_text("用法: `/deladmin <用户ID>`", parse_mode="Markdown")
        return

    target = context.args[0].strip()
    chat_data = await data_manager.get_chat_data(chat_id_str)
    admins = chat_data.setdefault("admins", [])

    if target not in admins:
        await update.message.reply_text("❌ 该用户不是机器人管理员")
        return

    admins.remove(target)
    await data_manager.update_chat_data(chat_id_str, chat_data)
    await data_manager.force_save()

    await update.message.reply_text(f"✅ 已删除管理员\n👤 ID: `{target}`", parse_mode="Markdown")


async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return

    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    owner = await get_group_owner(context, int(chat_id_str))

    text = "📋 **管理员列表**\n\n"
    if owner:
        text += f"👑 群主: `{owner}`\n\n"
    custom = chat_data.get("admins", [])
    text += f"🔧 机器人管理员 ({len(custom)}人):\n"
    if custom:
        for i, aid in enumerate(custom, 1):
            text += f"{i}. `{aid}`\n"
    else:
        text += "暂无\n"
    await update.message.reply_text(text, parse_mode="Markdown")


async def deluser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    if not context.args:
        await update.message.reply_text("用法: /deluser <用户ID 或 @用户名>")
        return
    chat_id_str = str(update.effective_chat.id)
    target = context.args[0].strip()
    chat_data = await data_manager.get_chat_data(chat_id_str)
    target_id = None
    if target.startswith('@'):
        name_search = target[1:].lower()
        for uid, name in chat_data["registered"].items():
            if name.lower() == name_search:
                target_id = uid
                break
    elif target in chat_data["registered"]:
        target_id = target

    if target_id and target_id in chat_data["registered"]:
        name = chat_data["registered"].pop(target_id)
        chat_data["users"].pop(target_id, None)
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await data_manager.force_save()
        await update.message.reply_text(f"✅ 已删除用户：**{name}**", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ 未找到该用户")


async def delete_record(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    if not context.args:
        await update.message.reply_text("用法: `/del YYYY-MM-DD`", parse_mode="Markdown")
        return
    date_to_del = context.args[0].strip()
    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    count = 0
    for user_id, user_info in list(chat_data.get("users", {}).items()):
        records_dict = user_info.get("records", {})
        if date_to_del in records_dict:
            del records_dict[date_to_del]
            count += 1
    if count == 0:
        await update.message.reply_text(f"ℹ️ 日期 **{date_to_del}** 没有记录", parse_mode="Markdown")
        return
    await data_manager.update_chat_data(chat_id_str, chat_data)
    await data_manager.force_save()
    await update.message.reply_text(f"✅ 已删除 **{date_to_del}** 的所有打卡记录（影响 {count} 人）", parse_mode="Markdown")


# ================== 报表命令 ==================
async def todayexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    today = get_attendance_date(beijing_now())
    chat_data = await data_manager.get_chat_data(chat_id_str)
    rows = build_daily_report_rows(chat_data, today)

    filename = f"全群打卡_{today}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
            "第一班迟到","第二班迟到","休息次数","总休息分钟",
            "工作原因休息次数","工作原因总休息分钟","状态"]
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    df.to_excel(filepath, index=False)

    with open(filepath, 'rb') as f:
        await update.message.reply_document(f, filename=filename, caption=f"✅ {today} 全群打卡报表")


async def monthexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    month = beijing_now().strftime("%Y-%m")
    chat_data = await data_manager.get_chat_data(chat_id_str)
    rows = build_month_report_rows(chat_data, month)

    filename = f"全群打卡_{month}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
            "第一班迟到","第二班迟到","休息次数","总休息分钟",
            "工作原因休息次数","工作原因总休息分钟","状态"]
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    df.to_excel(filepath, index=False)

    with open(filepath, 'rb') as f:
        await update.message.reply_document(f, filename=filename, caption=f"✅ {month} 月报表")


async def absent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update, context):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    today = get_attendance_date(beijing_now())
    chat_data = await data_manager.get_chat_data(chat_id_str)
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})

    incomplete = []
    for uid, name in registered.items():
        records = users.get(uid, {}).get("records", {}).get(today, [])
        done = {r["action"] for r in records if r.get("action") in "1234"}
        if done != {"1","2","3","4"}:
            incomplete.append(f"{name} → 已打: {','.join(sorted(done)) if done else '无'}")

    if not incomplete:
        await update.message.reply_text("🎉 今天所有人均已完成全部打卡！")
    else:
        text = f"📋 **今日未完成打卡人员** ({len(incomplete)}/{len(registered)})\n\n"
        text += "\n".join(f"{i+1}. {item}" for i, item in enumerate(incomplete))
        await update.message.reply_text(text, parse_mode="Markdown")

# ================== 自动日报 ==================
async def send_daily_report(context: ContextTypes.DEFAULT_TYPE):
    now = beijing_now()
    report_date = get_report_date_for_daily()
    
    print(f"🕒 自动日报触发 | 时间: {now} | 报表日期: {report_date}")
    
    cleanup_old_excels()

    # 全局一次性加载，避免多次 aload
    all_data = await data_manager.aload(force=True)
    sent_count = 0
    
    for chat_id_str, chat_data in list(all_data.items()):
        if not chat_id_str.startswith('-'):
            continue
            
        chat_id = int(chat_id_str)
        recipients = set()
        
        owner = await get_group_owner(context, chat_id)
        if owner:
            recipients.add(int(owner))
        recipients.update(int(uid) for uid in chat_data.get("admins", []))

        if not recipients:
            continue

        try:
            # 使用已经加载的数据，避免再次 get_chat_data
            rows = build_daily_report_rows(chat_data, report_date)
            
            filename = f"全群打卡日报_{report_date}.xlsx"
            filepath = os.path.join(EXCEL_FOLDER, filename)

            cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
                    "第一班迟到","第二班迟到","休息次数","总休息分钟",
                    "工作原因休息次数","工作原因总休息分钟","状态"]
            
            df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
            df.to_excel(filepath, index=False)

            caption = f"📊 **{report_date} 全群日报**"

            success = 0
            for rid in recipients:
                try:
                    with open(filepath, 'rb') as f:
                        await context.bot.send_document(rid, f, filename=filename, caption=caption, parse_mode="Markdown")
                    success += 1
                except Exception as e:
                    print(f"发送给 {rid} 失败: {e}")
            
            print(f"群 {chat_id} 日报发送完成 → {success}/{len(recipients)}")
            sent_count += 1
            
        except Exception as e:
            print(f"群 {chat_id} 处理异常: {e}")
    
    print(f"🎉 自动日报任务完成，共处理 {sent_count} 个群组")

# ================== 其他命令 ==================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "飞机的代号确定下来了就不要再改了，否则会打卡记录失败\n\n"
        "第一班上班打1，下班打2。第二班上班打3，下班打4，离开工位休息打5，回来打6（离开工位没打卡1次50，不论任何原因）\n\n"
        "上下班打卡的，迟到早退相同，10分钟内扣50，1小时内扣100，1小时外按旷工扣200。上班根据机器人的打卡时间，超过1秒也算迟到。漏打卡每次100\n"
        "⚠️严禁互相打卡与飞机定时发送。互相打卡两个人各扣300，定时发送扣600⚠️\n"
        "下班没打卡的不管是加班聊客户或者其他原因没打卡的一律算漏打卡。（下班打卡有效时间1小时）\n\n"
        "私聊机器人发送 /myrecord 可查询个人打卡记录\n"
    )


async def registered_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    registered = chat_data.get("registered", {})
    if not registered:
        await update.message.reply_text("📋 本群暂无注册人员。")
        return
    text = f"📋 **本群已注册人员**（{len(registered)}人）\n\n"
    for i, (uid, name) in enumerate(registered.items(), 1):
        text += f"{i}. {name} (`{uid}`)\n"
    await update.message.reply_text(text, parse_mode="Markdown")


async def myrecord(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """查询个人打卡记录 - 仅支持私聊"""
    if update.effective_chat.type != "private":
        await update.message.reply_text("⚠️ 此命令仅支持在 **私聊** 中使用")
        return

    user_id = str(update.effective_user.id)
    user_name = update.effective_user.full_name
    
    data = await data_manager.aload()
    
    text = f"📋 **{user_name}** 的打卡记录\n\n"
    found = False
    total_records = 0

    # 只遍历群组数据（负数ID）
    for chat_id, cdata in data.items():
        if not chat_id.startswith('-'):  # 跳过私聊残留
            continue
            
        urec = cdata.get("users", {}).get(user_id, {}).get("records", {})
        if not urec:
            continue

        found = True
        group_name = cdata.get("title", f"群 {chat_id}")  # 如果有群名称更好
        text += f"**📍 {group_name}**\n"
        
        # 按日期倒序，最多显示最近15天
        for date in sorted(urec.keys(), reverse=True)[:15]:
            recs = urec[date]
            if not recs:
                continue
                
            text += f"**{date}**\n"
            for r in recs:
                late = f"（迟到{r.get('late_seconds', 0)}秒）" if r.get("late_seconds") else ""
                display = r.get('display', r.get('action', '未知'))
                text += f"• {display}{late} {r.get('time', '')}\n"
            
            total_records += len(recs)
            text += "\n"

    if not found:
        text += "📭 暂无打卡记录\n\n您可以在群聊中发送 1-8 进行打卡。"
    else:
        text += f"共显示最近 {total_records} 条记录（最多显示15天）"

    await update.message.reply_text(text, parse_mode="Markdown")


async def today_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = beijing_now()
    att_date = get_attendance_date(now)
    report_date = get_report_date_for_daily()
    await update.message.reply_text(
        f"🕒 当前北京时间：**{now.strftime('%Y-%m-%d %H:%M:%S')}**\n"
        f"📅 当前考勤日期：**{att_date}**\n"
        f"📊 今日03:00将发送的日报日期：**{report_date}**",
        parse_mode="Markdown"
    )

# ================== 全局错误处理器 ==================
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """捕获所有错误"""
    error = context.error
    print(f"❌ 【全局错误】 {type(error).__name__}: {error}")

    # === 关键修复：使用字符串判断或直接导入 ===
    error_str = str(type(error).__name__).lower()
    
    if "networkerror" in error_str or "readerror" in error_str or "timeout" in error_str:
        print("🌐 检测到 Telegram 网络错误（Bad Gateway / ReadError / Timeout），Bot 将继续运行...")
        await asyncio.sleep(5)
        return

    if "telegramerror" in error_str:
        print("⚠️ Telegram API 错误，Bot 继续运行...")
        return

    # 其他严重错误打印完整堆栈
    import traceback
    print("🔥 严重错误:")
    print(traceback.format_exc())
    
    # 可选：通知群主或管理员（把 YOUR_ADMIN_ID 改成你的ID）
    # try:
    #     await context.bot.send_message(YOUR_ADMIN_ID, f"🚨 机器人发生错误:\n{error}")
    # except:
    #     pass

# ================== 主程序 ==================
def main():
    global data_manager
    data_manager = DataManager()
    
    print("📦 DataManager 初始化完成")
    data_manager.load(force=True)
    
    app = Application.builder() \
        .token(TOKEN) \
        .defaults(None) \
        .build()

    # ================== 【新增】注册全局错误处理器 ==================
    app.add_error_handler(error_handler)
    # ============================================================

    jq: JobQueue = app.job_queue
    
    beijing_tz = ZoneInfo("Asia/Shanghai")
    
    daily_time = datetime_time(3, 0, 0, tzinfo=beijing_tz)
    cleanup_time = datetime_time(3, 10, 0, tzinfo=beijing_tz)
    
    jq.run_daily(send_daily_report, daily_time)
    jq.run_daily(data_manager.cleanup_old_data, cleanup_time)

    print(f"⏰ 已设置自动任务：")
    print(f"   • 数据清理 → 北京时间 {cleanup_time}")
    print(f"   • 自动日报 → 北京时间 {daily_time}")

    handlers = [
        CommandHandler("start", start),
        CommandHandler("jihuo", activate_group),      # 新增激活指令
        CommandHandler("register", auto_register),
        CommandHandler("registered", registered_list),
        CommandHandler("myrecord", myrecord),
        CommandHandler("addadmin", add_admin),
        CommandHandler("deladmin", del_admin),
        CommandHandler("adminlist", adminlist),
        CommandHandler("deluser", deluser),
        CommandHandler("del", delete_record),
        CommandHandler("todayexcel", todayexcel),
        CommandHandler("monthexcel", monthexcel),
        CommandHandler("absent", absent),
        CommandHandler("today", today_cmd),
    ]

    for h in handlers:
        app.add_handler(h)

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_daka))

    print("🚀 打卡机器人已完全启动（啊财的机器人  7.13 ）")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
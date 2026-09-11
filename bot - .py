import json
import os
import asyncio
from datetime import datetime, timedelta, time as datetime_time
import pandas as pd
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

# ================== 新日期边界逻辑：04:00 为分界点 ==================
def get_attendance_date(now=None):
    """获取考勤所属日期（04:00 之后为当天）"""
    if now is None:
        now = beijing_now()
    if now.hour < 4:
        return beijing_date_str(now - timedelta(days=1))
    return beijing_date_str(now)


def get_record_date(shift: str, now=None) -> str:
    """根据打卡类型获取正确的记录日期 - 已适配04:00分界"""
    if now is None:
        now = beijing_now()
    
    base_date = get_attendance_date(now)
    
    if shift == "4" and now.hour < 4:
        return beijing_date_str(now - timedelta(days=1))
    
    return base_date


def get_previous_attendance_date(now=None) -> str:
    if now is None:
        now = beijing_now()
    return beijing_date_str(now - timedelta(days=1))


def get_report_date_for_daily() -> str:
    return get_previous_attendance_date(beijing_now())


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
    
    if datetime_time(12, 0) <= current_time < datetime_time(18, 0):
        return True, ""
    if datetime_time(19, 0) <= current_time or current_time < datetime_time(23, 59):
        return True, ""
    
    return False, "⚠️ 休息/暂离（5或7）只能在以下工作时段打卡：\n• 第一班 12:00-18:00\n• 第二班 19:00-00:00"


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


# ================== 状态判断工具函数 ==================
def is_currently_on_duty(records: list) -> bool:
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
            if now.hour < 4:  
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


# ================== DataManager ==================
class DataManager:
    def __init__(self):
        self._data: dict = {}
        self._last_mtime = 0
        self._last_save = 0
        self._dirty = False
        self._global_lock = asyncio.Lock()
        self._chat_locks: dict[str, asyncio.Lock] = {}
        self._save_task = None
        self._migrated = False

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
                    print(f"📥 数据已从磁盘加载 | 群组: {len(self._data)}")
                except Exception as e:
                    print(f"❌ 加载数据失败: {e}")
                    self._data = {}
            else:
                self._data = {}
            
            self._last_mtime = current_mtime
            self._dirty = False
            
            if not self._migrated:
                self._migrate_historical_data()
                self._migrated = True
        return self._data

    def _migrate_historical_data(self):
        print("🔄 开始执行历史数据日期迁移...")
        migrated_count = 0
        for chat_id, chat_data in self._data.items():
            users = chat_data.get("users", {})
            for user_id, user_info in users.items():
                records = user_info.get("records", {})
                new_records: dict[str, list] = {}
                for old_date, rec_list in list(records.items()):
                    for rec in rec_list:
                        action = rec.get("action")
                        time_str = rec.get("time", "00:00:00")
                        try:
                            rec_time = datetime.strptime(time_str, "%H:%M:%S").time()
                            dummy_dt = datetime.strptime(old_date, "%Y-%m-%d").replace(
                                hour=rec_time.hour, minute=rec_time.minute, 
                                second=rec_time.second, tzinfo=TZ
                            )
                            new_date = get_record_date(action, dummy_dt)
                            if new_date not in new_records:
                                new_records[new_date] = []
                            if not any(r.get("time") == rec.get("time") and r.get("action") == action for r in new_records.get(new_date, [])):
                                new_records[new_date].append(rec.copy())
                                if new_date != old_date:
                                    migrated_count += 1
                        except Exception:
                            if old_date not in new_records:
                                new_records[old_date] = []
                            new_records[old_date].append(rec.copy())
                user_info["records"] = new_records
        if migrated_count > 0:
            self._dirty = True
            print(f"✅ 历史数据迁移完成，共调整 {migrated_count} 条记录")

    async def aload(self, force: bool = False) -> dict:
        return await asyncio.to_thread(self.load, force)

    async def save(self, immediate: bool = False):
        async with self._global_lock:
            if not self._dirty and not immediate:
                return
            try:
                temp_file = DATA_FILE + ".tmp"
                backup_file = DATA_FILE + ".bak"
                with open(temp_file, "w", encoding="utf-8") as f:
                    json.dump(self._data, f, ensure_ascii=False, indent=2)
                if os.path.exists(DATA_FILE):
                    shutil.copy2(DATA_FILE, backup_file)
                os.replace(temp_file, DATA_FILE)
                self._last_mtime = self._file_mtime()
                self._last_save = time_module.time()
                self._dirty = False
                print(f"💾 数据已安全保存 | 群组: {len(self._data)}")
            except Exception as e:
                print(f"❌ 保存失败: {e}")

    async def _delayed_save(self):
        await asyncio.sleep(3)
        await self.save()

    async def get_chat_data(self, chat_id: str):
        async with self._get_chat_lock(chat_id):
            await self.aload()
            return self._data.setdefault(chat_id, {
                "registered": {},
                "users": {},
                "admins": [],
                "activated": False  # 默认未激活
            })

    async def update_chat_data(self, chat_id: str, chat_data: dict):
        async with self._get_chat_lock(chat_id):
            await self.aload()
            self._data[chat_id] = chat_data
            self._dirty = True
            if not self._save_task or self._save_task.done():
                self._save_task = asyncio.create_task(self._delayed_save())

    async def force_save(self):
        await self.save(immediate=True)

    async def cleanup_old_data(self, context: ContextTypes.DEFAULT_TYPE = None):
        """清理90天前的旧记录"""
        async with self._global_lock:
            await self.aload(force=True)
            cutoff = (beijing_now() - timedelta(days=90)).strftime("%Y-%m-%d")
            cleaned = 0
            
            for chat_id in list(self._data.keys()):
                for user_id in list(self._data[chat_id].get("users", {}).keys()):
                    records = self._data[chat_id]["users"][user_id].get("records", {})
                    for d in list(records.keys()):
                        if d < cutoff:
                            del records[d]
                            cleaned += 1
            
            if cleaned > 0:
                self._dirty = True
                print(f"🧹 已清理 {cleaned} 条旧记录")
                await self.force_save()
            else:
                print("🧹 没有需要清理的旧记录")


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
    # 如果没有传入chat_data，从全局读取
    return data_manager._data.get(chat_id, {}).get("activated", False)


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
    """日报 - 单日报表"""
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})
    rows = []
    
    for user_id, user_name in registered.items():
        user_info = users.get(user_id, {"name": user_name, "records": {}})
        records = user_info.get("records", {}).get(report_date, [])

        shifts = {r.get("action"): r for r in records if r.get("action") in {"1", "2", "3", "4"}}

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
    """月报表 - 按天展开"""
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})
    rows = []
    
    for user_id, user_name in registered.items():
        user_records = users.get(user_id, {}).get("records", {})
        
        for date, records in user_records.items():
            if not date.startswith(month):
                continue
                
            shifts = {r.get("action"): r for r in records if r.get("action") in {"1", "2", "3", "4"}}

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


# ================== 核心打卡函数 ==================
async def daka(update: Update, context: ContextTypes.DEFAULT_TYPE, shift: str):
    chat_id_str = str(update.effective_chat.id)
    user = update.effective_user
    user_id = str(user.id)

    # ================== 激活检查 ==================
    if update.effective_chat.type != "private":
        chat_data_temp = await data_manager.get_chat_data(chat_id_str)
        if not chat_data_temp.get("activated", False):
            await update.message.reply_text(
                "⚠️ **本群尚未激活**\n\n"
                "此机器人需要密码才能激活使用。\n"
                "请联系机器人管理员获取激活密码。",
                parse_mode="Markdown"
            )
            return

    await auto_register(update, context)

    now = beijing_now()
    time_str = now.strftime("%H:%M:%S")
    date_str = get_record_date(shift, now)

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

    chat_data = await data_manager.get_chat_data(chat_id_str)
    user_data = chat_data["users"].setdefault(user_id, {"name": user.full_name, "records": {}})
    records: list = user_data["records"].setdefault(date_str, [])

    # ================== 构建模拟记录 ==================
    simulated_records = [r.copy() for r in records]
    action_info = ACTIONS.get(shift, {"name": shift, "type": "unknown"})

    check_records = [r.copy() for r in records]

    if shift in ["1", "2", "3", "4"]:
        simulated_records.append({"action": shift, "type": "work"})
    elif shift == "5":
        simulated_records.append({"type": "rest_start"})
    elif shift == "6":
        simulated_records.append({"action": shift, "type": "rest_end"})
    elif shift == "7":
        simulated_records.append({"type": "work_rest_start"})
    elif shift == "8":
        simulated_records.append({"action": shift, "type": "work_rest_end"})

    # ================== 状态判断 ==================
    is_resting = any(
        r.get("type") == "rest_start" and "rest_minutes" not in r 
        for r in records   # 改用 records 而非 check_records
    )
    is_work_resting = any(
        r.get("type") == "work_rest_start" and "rest_minutes" not in r 
        for r in records
    )

    is_on_duty = is_currently_on_duty(simulated_records)
    has_started = has_started_work_today(check_records)

    # ================== 重复打卡检查 ==================
    if shift in ["1", "2", "3", "4"]:
        if any(r.get("action") == shift for r in records):
            await update.message.reply_text(f"⚠️ {date_str} 已打过 {ACTIONS[shift]['name']}")
            return

    # ================== 业务规则检查 ==================
    if shift == "6" and not is_resting:
        await update.message.reply_text("⚠️ 请先输入5开始休息")
        return

    if shift == "8" and not is_work_resting:
        await update.message.reply_text("⚠️ 请先输入7工作原因暂离")
        return

    # ================== 必须结束休息才能下班 ==================
    if shift in ["2", "4"]:
        open_rest = any(
            r.get("type") == "rest_start" and "rest_minutes" not in r 
            for r in records
        )
        open_work_rest = any(
            r.get("type") == "work_rest_start" and "rest_minutes" not in r 
            for r in records
        )
        
        if open_rest or open_work_rest:
            rest_type = "休息" if open_rest else "工作原因暂离"
            await update.message.reply_text(
                f"⚠️ **您目前处于「{rest_type}」状态**\n\n"
                "请先回复 **6** 结束休息（或 **8** 结束暂离），\n"
                "再打下班卡（2 或 4）。"
            )
            return

    # ================== 修改点3：移除必须上班才能打5/7的限制 ==================
    if shift in ["5", "7"]:
        if is_resting or is_work_resting:
            await update.message.reply_text("⏳ 当前正在休息中，请先结束再开始新休息")
            return
        # 不再检查 is_on_duty 和 has_started

    # ================== 执行实际打卡 ==================
    late_seconds, late_txt = get_late_minutes(action_info.get("time"), shift, now)
    display = action_info["name"]
    final_display = display

    if shift in ["6", "8"]:
        target_type = "rest_start" if shift == "6" else "work_rest_start"
        matched = False
        for r in reversed(records):
            if r.get("type") == target_type and "rest_minutes" not in r:
                rest_min = calculate_rest_duration(r["time"], time_str)
                final_display = f"{action_info['name']}（{rest_min}分钟）"
                
                records.append({
                    "time": time_str,
                    "action": shift,
                    "display": final_display,
                    "rest_minutes": rest_min,
                    "type": action_info.get("type")
                })
                r["rest_minutes"] = rest_min
                matched = True

                if shift == "6":
                    job_name = r.get("rest_job_name") or f"rest_timeout_{chat_id_str}_{user_id}_{r.get('time')}"
                    
                    removed = 0
                    for job in list(context.job_queue.get_jobs_by_name(job_name)):
                        job.schedule_removal()
                        removed += 1
                    
                    if removed == 0:
                        prefix = f"rest_timeout_{chat_id_str}_{user_id}_"
                        for job in context.job_queue.jobs():
                            if job.name and job.name.startswith(prefix):
                                job.schedule_removal()
                                removed += 1
                    
                    print(f"🛑 已取消休息超时任务: {job_name} (移除 {removed} 个)")
                
                break
        
        if not matched:
            await update.message.reply_text(f"⚠️ 未找到对应的开始记录，无法结束{action_info['name']}")
            return

    else:
        record_entry = {
            "time": time_str,
            "action": shift,
            "display": display,
            "type": action_info.get("type")
        }
        
        if late_seconds > 0 and shift in ["1", "3"]:
            record_entry["late_seconds"] = late_seconds
            record_entry["late_display"] = late_txt
            record_entry["display"] = f"{display}{late_txt}"
            final_display = record_entry["display"]

        records.append(record_entry)

        if shift == "5":
            job_name = f"rest_timeout_{chat_id_str}_{user_id}_{time_str}"
            record_entry["rest_job_name"] = job_name

            context.job_queue.run_once(
                callback=check_rest_timeout,
                when=3600,
                data={
                    "chat_id": int(chat_id_str), 
                    "user_id": user_id, 
                    "start_time": time_str,
                    "job_name": job_name
                },
                name=job_name
            )

    await data_manager.update_chat_data(chat_id_str, chat_data)

    emoji = "⚠️" if late_seconds > 0 else "✅"
    await update.message.reply_text(
        f"{emoji} **{user.full_name}** {final_display}\n日期：{date_str}\n时间：{time_str}",
        parse_mode="Markdown"
    )


# ================== 消息处理 ==================
async def text_daka(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    chat_id_str = str(update.effective_chat.id)
    user_id = str(update.effective_user.id)
    name = update.effective_user.full_name

    chat_data = await data_manager.get_chat_data(chat_id_str)
    if user_id not in chat_data["registered"]:
        chat_data["registered"][user_id] = name
        chat_data["users"].setdefault(user_id, {"name": name, "records": {}})
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await update.message.reply_text(f"✅ **{name}** 自动注册成功！", parse_mode="Markdown")


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
    
    print(f"🕒 自动日报任务触发 | 北京时间: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📊 准备发送的报表日期: {report_date}")
    
    cleanup_old_excels()

    all_data = await data_manager.aload(force=True)
    sent_count = 0
    
    for chat_id_str, chat_data in all_data.items():
        chat_id = int(chat_id_str)
        recipients = set()
        
        owner = await get_group_owner(context, chat_id)
        if owner:
            recipients.add(int(owner))
        recipients.update(int(uid) for uid in chat_data.get("admins", []))

        if not recipients:
            print(f"⚠️ 群 {chat_id} 没有收件人")
            continue

        try:
            fresh_chat_data = await data_manager.get_chat_data(chat_id_str)
            rows = build_daily_report_rows(fresh_chat_data, report_date)
            
            filename = f"全群打卡日报_{report_date}.xlsx"
            filepath = os.path.join(EXCEL_FOLDER, filename)

            cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
                    "第一班迟到","第二班迟到","休息次数","总休息分钟",
                    "工作原因休息次数","工作原因总休息分钟","状态"]
            
            df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
            df.to_excel(filepath, index=False)

            caption = f"📊 **{report_date} 全群日报**（02:00~次日02:00）"

            success = 0
            for rid in recipients:
                try:
                    with open(filepath, 'rb') as f:
                        await context.bot.send_document(
                            rid, 
                            f, 
                            filename=filename, 
                            caption=caption, 
                            parse_mode="Markdown"
                        )
                    success += 1
                except Exception as e:
                    print(f"❌ 发送给 {rid} 失败: {e}")
            
            print(f"✅ 群 {chat_id} 日报发送完成 → {success}/{len(recipients)} 人接收")
            sent_count += 1
            
        except Exception as e:
            print(f"❌ 群 {chat_id} 生成/发送日报异常: {e}")
    
    print(f"🎉 自动日报任务全部完成，共处理 {sent_count} 个群组")


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
    if update.effective_chat.type != "private":
        await update.message.reply_text("此命令仅支持私聊使用")
        return
    user_id = str(update.effective_user.id)
    data = await data_manager.aload()
    text = f"📋 **{update.effective_user.full_name}** 打卡记录\n\n"
    found = False
    for chat_id, cdata in data.items():
        urec = cdata.get("users", {}).get(user_id, {}).get("records", {})
        if not urec: continue
        found = True
        text += f"**群 {chat_id}**\n"
        for date in sorted(urec.keys(), reverse=True)[:15]:
            recs = urec[date]
            if not recs: continue
            text += f"**{date}**\n"
            for r in recs:
                late = f"（迟到{r.get('late_seconds',0)}秒）" if r.get("late_seconds") else ""
                text += f"• {r.get('display')}{late} {r['time']}\n"
            text += "\n"
    await update.message.reply_text(text if found else "暂无记录", parse_mode="Markdown")


async def today_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = beijing_now()
    att_date = get_attendance_date(now)
    report_date = get_report_date_for_daily()
    await update.message.reply_text(
        f"🕒 当前北京时间：**{now.strftime('%Y-%m-%d %H:%M:%S')}**\n"
        f"📅 当前考勤日期：**{att_date}**\n"
        f"📊 今日04:30将发送的日报日期：**{report_date}**",
        parse_mode="Markdown"
    )


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
    
    jq: JobQueue = app.job_queue
    
    beijing_tz = ZoneInfo("Asia/Shanghai")
    
    daily_time = datetime_time(4, 30, 0, tzinfo=beijing_tz)
    cleanup_time = datetime_time(4, 40, 0, tzinfo=beijing_tz)
    
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

    print("🚀 打卡机器人已完全启动（啊原的机器人  6.22 ）")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
# -*- coding: utf-8 -*-
import webview
import get_xls_data
import threading
import time
import re
from datetime import datetime, timedelta
import os
import sys

def get_resource_path(relative_path):
    """获取资源文件的绝对路径（用于打包进exe的资源）"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_data_path(relative_path):
    """获取外部数据文件的绝对路径（用户自己放的数据）"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class Api:

    def __init__(self):
        self.real_time_data = {}
        self.history_data = {}
        self.merged_data = {}
        self.concept_data = {}
        self.stock_tracking = {}
        self.industry_data = get_xls_data.get_code_industry()
        self.auto_update_running = False
        self.update_thread = None
        self.last_update_time = None
        self.data_source_info = ''
        print('API 初始化完成')

    def calculate_workdays(self, start_date, end_date):
        """计算两个日期之间的工作日天数（排除周六周日）"""
        days = 0
        current = start_date + timedelta(days=1)
        while current <= end_date:
            if current.weekday() < 5:
                days += 1
            current += timedelta(days=1)
        return days

    def is_limit_up(self, stock_code, change_pct):
        """判断是否涨停（严格标准）

        Args:
            stock_code: 股票代码
            change_pct: 涨幅百分比

        Returns:
            bool: 是否涨停
        """
        if stock_code.startswith('3') or stock_code.startswith('68'):
            return change_pct >= 19.8
        return change_pct >= 9.8

    def is_limit_up_loose(self, stock_code, change_pct):
        """判断是否涨停（宽松标准：所有股票统一9.8%）

        Args:
            stock_code: 股票代码
            change_pct: 涨幅百分比

        Returns:
            bool: 是否涨停
        """
        return change_pct >= 9.8

    def get_concept_data(self, strat_index=3, count=21):
        if not self.concept_data:
            print('正在加载概念数据...')
            self.concept_data = get_xls_data.get_folder_data(strat_index=strat_index, count=count)
            print(f'概念数据加载完成：{len(self.concept_data)} 天')
            debug_codes = ['605188', '002337', '600262']
            print('\n【概念数据检查】')
            for code in debug_codes:
                found = False
                for date_str, stocks_list in self.concept_data.items():
                    for stock_data in stocks_list:
                        if stock_data[0] == code:
                            found = True
                            print(f'  {code} 在概念数据中 (日期: {date_str}, 名称: {stock_data[1]})')
                            break
                    if found:
                        break
                if not found:
                    print(f'  {code} 不在概念数据中')
        return self.concept_data

    def classify_priority_stocks(self, strat_index=3, count=21):
        """根据历史数据分类股票优先级
        Returns:
            dict: {'high_priority': [], 'normal_priority': []}
        """
        print('\n========== 开始股票优先级分类 ==========')
        if not self.history_data:
            print('历史数据为空，先获取历史数据...')
            self.get_history_data(strat_index=strat_index, count=count, show_progress=False)
        if not self.concept_data:
            auto_index, reason = get_xls_data.get_data_source_index()
            actual_index = strat_index if strat_index else auto_index
            self.concept_data = get_xls_data.get_folder_data(strat_index=actual_index, count=count)
        high_priority_stocks = []
        normal_priority_stocks = []
        all_stock_codes = set()
        for date, stocks_list in self.concept_data.items():
            for stock_data in stocks_list:
                stock_code = stock_data[0]
                all_stock_codes.add(stock_code)
        print(f'总股票数: {len(all_stock_codes)}')
        for stock_code in all_stock_codes:
            if stock_code not in self.history_data:
                normal_priority_stocks.append(stock_code)
            else:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                if not price_list or len(price_list) < 2:
                    normal_priority_stocks.append(stock_code)
                else:
                    closes = [p['收盘价'] for p in price_list]
                    sunny_days = 0
                    for i in range(len(closes) - 1, 0, -1):
                        if closes[i] > closes[i - 1]:
                            sunny_days += 1
                        else:
                            break
                    has_consecutive_limit_up = False
                    limit_up_indices = []
                    for i, price_data in enumerate(price_list):
                        change_pct = price_data.get('涨幅', 0)
                        if change_pct >= 9.8:
                            limit_up_indices.append(i)
                    if len(limit_up_indices) >= 2:
                        for i in range(len(limit_up_indices) - 1):
                            if limit_up_indices[i + 1] - limit_up_indices[i] == 1:
                                has_consecutive_limit_up = True
                                break
                    if sunny_days == 1 and has_consecutive_limit_up:
                        high_priority_stocks.append(stock_code)
                    else:
                        normal_priority_stocks.append(stock_code)
        print(f'高优先级股票（阳天数=1 + 有连续涨停）: {len(high_priority_stocks)}个')
        print(f'普通优先级股票: {len(normal_priority_stocks)}个')
        print('==================================================')
        return {'high_priority': high_priority_stocks, 'normal_priority': normal_priority_stocks}

    def get_real_time_data(self, strat_index=3, count=21, show_progress=True, priority_codes=None):

        def test_callback(current, total, msg):
            try:
                webview.windows[0].evaluate_js(f'updateProgress({current}, {total}, "{msg}")')
            except Exception as e:
                pass
            return None
        auto_index, reason = get_xls_data.get_data_source_index()
        actual_index = strat_index if strat_index else auto_index
        self.data_source_info = reason
        print(f'数据源选择: {reason}, 使用索引: {actual_index}')
        self.concept_data = get_xls_data.get_folder_data(strat_index=actual_index, count=count)
        self.get_history_data(strat_index=actual_index, count=count, show_progress=False)
        classification = self.classify_priority_stocks(strat_index=actual_index, count=count)
        high_priority_codes = classification['high_priority']
        result = get_xls_data.get_real_time_data(progress_callback=test_callback, strat_index=actual_index, count=count, show_progress=show_progress, top_priority_codes=priority_codes, high_priority_codes=high_priority_codes)
        self.real_time_data = result
        self.last_update_time = datetime.now().strftime('%H:%M:%S')
        self.check_breakthrough()
        debug_codes = ['605188', '002337', '600262']
        print('\n【实时数据检查】')
        for code in debug_codes:
            sh_code = f'sh{code}'
            sz_code = f'sz{code}'
            if sh_code in self.real_time_data:
                print(f'  {code} 在实时数据中 (sh{code})')
            elif sz_code in self.real_time_data:
                print(f'  {code} 在实时数据中 (sz{code})')
            else:
                print(f'  {code} 不在实时数据中')
        return {'概念数据': self.concept_data, '实时数据': result, '更新时间': self.last_update_time, '数据源': reason}

    def get_history_data(self, strat_index=3, count=21, show_progress=True):

        def history_callback(current, total, msg):
            try:
                webview.windows[0].evaluate_js(f'updateProgress({current}, {total}, "{msg}")')
            except Exception as e:
                pass
            return None
        auto_index, reason = get_xls_data.get_data_source_index()
        actual_index = strat_index if strat_index else auto_index
        result = get_xls_data.get_history_data(progress_callback=history_callback, strat_index=actual_index, count=count, show_progress=show_progress)
        self.history_data = result
        debug_codes = ['605188', '002337', '600262']
        print('\n【历史数据检查】')
        for code in debug_codes:
            if code in self.history_data:
                hist_data = self.history_data[code]
                price_list = hist_data.get('历史价格列表', [])
                print(f'  {code} 在历史数据中 (共{len(price_list)}天)')
            else:
                print(f'  {code} 不在历史数据中')
        self.merge_all_data(min_days=strat_index, max_days=count)
        return result

    def check_breakthrough(self):
        for prefix_code, real_data in self.real_time_data.items():
            stock_code = prefix_code[2:]
            try:
                current_price = float(real_data.get('现价', 0))
                current_zhangfu = float(real_data.get('涨幅', 0))
                if current_price <= 0:
                    continue
                max_30d = 0
                max_60d = 0
                if stock_code in self.history_data:
                    max_30d = float(self.history_data[stock_code].get('30日最高价', 0))
                    max_60d = float(self.history_data[stock_code].get('60日最高价', 0))
                if stock_code not in self.stock_tracking:
                    self.stock_tracking[stock_code] = {'初始价格': current_price, '当前最高': current_price, '当前最低': current_price, '突破新高次数': 0, '突破新低次数': 0, '曾经最高涨幅': current_zhangfu, '突破30日新高次数': 0, '上次是否超过30日': False, '已突破60日新高': False}
                    if max_30d > 0 and current_price > max_30d:
                        self.stock_tracking[stock_code]['突破30日新高次数'] = 1
                        self.stock_tracking[stock_code]['上次是否超过30日'] = True
                        print(f'股票 {stock_code} 首次突破30日新高: {current_price} > {max_30d}')
                    if max_60d > 0 and current_price > max_60d:
                        self.stock_tracking[stock_code]['已突破60日新高'] = True
                        print(f'股票 {stock_code} 突破60日新高: {current_price} > {max_60d}')
                else:
                    tracking = self.stock_tracking[stock_code]
                    if current_price > tracking['当前最高']:
                        tracking['突破新高次数'] += 1
                        tracking['当前最高'] = current_price
                        print(f"股票 {stock_code} 突破新高: {current_price}, 累计{tracking['突破新高次数']}次")
                    if current_price < tracking['当前最低']:
                        tracking['突破新低次数'] += 1
                        tracking['当前最低'] = current_price
                        print(f"股票 {stock_code} 跌破新低: {current_price}, 累计{tracking['突破新低次数']}次")
                    if current_zhangfu > tracking['曾经最高涨幅']:
                        tracking['曾经最高涨幅'] = current_zhangfu
                        print(f'股票 {stock_code} 创今日最高涨幅: {current_zhangfu}%')
                    if max_30d > 0:
                        currently_above_30d = current_price > max_30d
                        if currently_above_30d and (not tracking['上次是否超过30日']):
                            tracking['突破30日新高次数'] += 1
                            tracking['上次是否超过30日'] = True
                            print(f"股票 {stock_code} 突破30日新高: {current_price} > {max_30d}, 累计{tracking['突破30日新高次数']}次")
                        elif not currently_above_30d:
                            tracking['上次是否超过30日'] = False
                    if max_60d > 0:
                        if current_price > max_60d:
                            tracking['已突破60日新高'] = True
                        else:
                            tracking['已突破60日新高'] = False
            except:
                continue

    def analyze_limit_up_streak(self, concept_dates=None, use_loose=False):
        """从历史数据分析连续涨停（基于索引位置判断）

        Args:
            concept_dates: 要分析的日期范围（set），如果为None则分析所有历史数据
            use_loose: 是否使用宽松标准（True=统一9.8%，False=按板块区分）
        """
        if not self.history_data:
            return {}
        result = {}
        for stock_code, hist_data in self.history_data.items():
            price_list = hist_data.get('历史价格列表', [])
            if not price_list:
                continue
            limit_up_info = []
            for i, price_data in enumerate(price_list):
                date_str = price_data['日期']
                if concept_dates and date_str not in concept_dates:
                    continue
                is_zt = self.is_limit_up_loose(stock_code, price_data['涨幅']) if use_loose else self.is_limit_up(stock_code, price_data['涨幅'])
                if is_zt:
                    limit_up_info.append({'index': i, 'date': date_str, 'change': price_data['涨幅']})
            if not limit_up_info:
                continue
            max_streak = 0
            current_streak = 1
            last_streak_end_index = None
            if len(limit_up_info) >= 2:
                for i in range(1, len(limit_up_info)):
                    prev_index = limit_up_info[i - 1]['index']
                    curr_index = limit_up_info[i]['index']
                    if curr_index - prev_index == 1:
                        current_streak += 1
                        if current_streak >= 2:
                            last_streak_end_index = curr_index
                    else:
                        if current_streak >= 2 and max_streak < current_streak:
                            max_streak = current_streak
                        current_streak = 1
                if current_streak >= 2 and max_streak < current_streak:
                    max_streak = current_streak
                    last_streak_end_index = limit_up_info[-1]['index']
            elif len(limit_up_info) == 1:
                pass
            last_index_in_full_list = len(price_list) - 1
            last_limit_up_index = limit_up_info[-1]['index']
            if max_streak >= 2 and last_streak_end_index is not None:
                days_diff = last_index_in_full_list - last_streak_end_index
                last_date = price_list[last_streak_end_index]['日期']
            else:
                days_diff = last_index_in_full_list - last_limit_up_index
                last_date = price_list[last_limit_up_index]['日期']
            result[stock_code] = {'最大连续涨停数': max_streak, '最后涨停日期': last_date, '离最新日期天数': days_diff}
        return result

    def _count_single_day_segments(self, indices):
        """计算长度为1的涨停段个数（单日涨停）"""
        if not indices:
            return 0
        singles = 0
        run_len = 1
        for i in range(1, len(indices)):
            if indices[i] == indices[i - 1] + 1:
                run_len += 1
            else:
                if run_len == 1:
                    singles += 1
                run_len = 1
        if run_len == 1:
            singles += 1
        return singles

    def merge_all_data(self, min_days=3, max_days=21):
        """合并所有数据
        
        Args:
            min_days: 离涨停天数的最小值（用于计算总涨停数）
            max_days: 离涨停天数的最大值（用于计算总涨停数）
        """
        merged = {}
        concept_dates = set()
        for date_str in self.concept_data.keys():
            match = re.match('(\\d+)月(\\d+)', date_str)
            if match:
                month = int(match.group(1))
                day = int(match.group(2))
                year = 2025
                concept_dates.add(f'{year:04d}-{month:02d}-{day:02d}')
        # 连续涨停数基于全部60天历史数据，不受Excel日期范围限制
        limit_up_info_strict = self.analyze_limit_up_streak(None, use_loose=False)
        limit_up_info_loose = self.analyze_limit_up_streak(None, use_loose=True)
        # 调试：检查002792是否在数据中
        print(f'\n【数据检查】')
        print(f'  real_time_data 总数: {len(self.real_time_data)}')
        print(f'  history_data 总数: {len(self.history_data)}')
        print(f'  sz002792 在 real_time_data 中: {"sz002792" in self.real_time_data}')
        print(f'  002792 在 history_data 中: {"002792" in self.history_data}')
        if '002792' in self.history_data:
            hist = self.history_data['002792']
            plist = hist.get('历史价格列表', [])
            print(f'  002792 历史价格列表长度: {len(plist)}')
        for prefix_code, real_data in self.real_time_data.items():
            stock_code = prefix_code[2:]
            tracking = self.stock_tracking.get(stock_code, {})
            limit_info_strict = limit_up_info_strict.get(stock_code, {})
            limit_info_loose = limit_up_info_loose.get(stock_code, {})
            sunny_days = 0
            if stock_code in self.history_data:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                if price_list and len(price_list) >= 2:
                    closes = [p['收盘价'] for p in price_list]
                    yesterday_close_from_hist = hist_data.get('昨日收盘价', 0)
                    try:
                        realtime_price = float(real_data.get('现价', 0))
                        if realtime_price > 0:
                            if len(price_list) >= 2 and abs(price_list[-2]['收盘价'] - yesterday_close_from_hist) < 0.01:
                                closes[-1] = realtime_price
                            else:
                                closes.append(realtime_price)
                    except Exception:
                        pass
                    for i in range(len(closes) - 1, 0, -1):
                        current_close = closes[i]
                        prev_close = closes[i - 1]
                        if current_close > prev_close:
                            sunny_days += 1
                        else:
                            break
            # 计算前N天连续阳涨幅天数（跳过最后两天，从倒数第三天开始往前数）
            prev_positive_days = 0
            if stock_code in self.history_data:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                if price_list and len(price_list) >= 3:
                    for i in range(len(price_list) - 3, -1, -1):
                        change = price_list[i].get('涨幅', 0)
                        if change > 0:
                            prev_positive_days += 1
                        else:
                            break
            limit_up_count_strict = 0
            limit_up_count_loose = 0
            single_count_strict = 0
            single_count_loose = 0
            total_all_limit_days_strict = 0
            total_all_limit_days_loose = 0
            # 统计全部60天的涨停天数，不依赖concept_dates
            if stock_code in self.history_data:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                limit_up_indices_strict = []
                for i, price_data in enumerate(price_list):
                    is_zt_strict = self.is_limit_up(stock_code, price_data['涨幅'])
                    # 统计全部60天内的涨停，不限制日期范围
                    if is_zt_strict:
                        limit_up_indices_strict.append(i)
                if limit_up_indices_strict:
                    limit_up_count_strict = 1
                    for i in range(1, len(limit_up_indices_strict)):
                        if limit_up_indices_strict[i] - limit_up_indices_strict[i - 1] > 1:
                            limit_up_count_strict += 1
                limit_up_indices_loose = []
                for i, price_data in enumerate(price_list):
                    is_zt_loose = self.is_limit_up_loose(stock_code, price_data['涨幅'])
                    # 统计全部60天内的涨停，不限制日期范围
                    if is_zt_loose:
                        limit_up_indices_loose.append(i)
                if limit_up_indices_loose:
                    limit_up_count_loose = 1
                    for i in range(1, len(limit_up_indices_loose)):
                        if limit_up_indices_loose[i] - limit_up_indices_loose[i - 1] > 1:
                            limit_up_count_loose += 1
                single_count_strict = self._count_single_day_segments(limit_up_indices_strict)
                single_count_loose = self._count_single_day_segments(limit_up_indices_loose)
                total_all_limit_days_strict = len(limit_up_indices_strict)
                total_all_limit_days_loose = len(limit_up_indices_loose)
                # 调试：打印002792的涨停统计
                if stock_code == '002792':
                    print(f'\n【002792 涨停统计调试】')
                    print(f'  history_data中存在: True')
                    print(f'  price_list长度: {len(price_list)}')
                    print(f'  涨停日索引(严格): {limit_up_indices_strict}')
                    print(f'  全部涨停天数(严格): {total_all_limit_days_strict}')
                    print(f'  连续涨停数(严格): {limit_info_strict.get("最大连续涨停数", 0)}')
                    print(f'  涨停日详情:')
                    for idx in limit_up_indices_strict:
                        p = price_list[idx]
                        print(f'    索引{idx}: {p["日期"]} 涨幅{p["涨幅"]}%')
            total_limit_up_days_in_range_strict = 0
            total_limit_up_days_in_range_loose = 0
            if stock_code in self.history_data and concept_dates:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                consecutive_limit_up_strict = limit_info_strict.get('最大连续涨停数', 0)
                consecutive_limit_up_loose = limit_info_loose.get('最大连续涨停数', 0)
                days_from_limit_strict = limit_info_strict.get('离最新日期天数', '无涨停')
                days_from_limit_loose = limit_info_loose.get('离最新日期天数', '无涨停')
                should_count_strict = False
                should_count_loose = False
                if consecutive_limit_up_strict >= 2 and days_from_limit_strict != '无涨停':
                    try:
                        days_strict = int(days_from_limit_strict)
                        if min_days <= days_strict <= max_days:
                            should_count_strict = True
                    except (ValueError, TypeError):
                        pass
                if consecutive_limit_up_loose >= 2 and days_from_limit_loose != '无涨停':
                    try:
                        days_loose = int(days_from_limit_loose)
                        if min_days <= days_loose <= max_days:
                            should_count_loose = True
                    except (ValueError, TypeError):
                        pass
                total_all_limit_up_days_strict = 0
                total_all_limit_up_days_loose = 0
                if should_count_strict:
                    for i, price_data in enumerate(price_list):
                        date_in_concept = price_data['日期'] in concept_dates
                        is_zt_strict = self.is_limit_up(stock_code, price_data['涨幅'])
                        if date_in_concept and is_zt_strict:
                            total_all_limit_up_days_strict += 1
                    total_limit_up_days_in_range_strict = max(0, total_all_limit_up_days_strict - consecutive_limit_up_strict)
                if should_count_loose:
                    for i, price_data in enumerate(price_list):
                        date_in_concept = price_data['日期'] in concept_dates
                        is_zt_loose = self.is_limit_up_loose(stock_code, price_data['涨幅'])
                        if date_in_concept and is_zt_loose:
                            total_all_limit_up_days_loose += 1
                    total_limit_up_days_in_range_loose = max(0, total_all_limit_up_days_loose - consecutive_limit_up_loose)
            time_info = get_xls_data.get_current_time_info()
            weekday = time_info['星期']
            hour = time_info['小时']
            minute = time_info['分钟']
            is_trading_time = weekday < 5 and (hour == 9 and minute >= 15 or 9 < hour < 15 or (hour == 15 and minute == 0))
            if is_trading_time:
                current_change = real_data.get('涨幅', '')
            elif stock_code in self.history_data:
                hist_data = self.history_data[stock_code]
                price_list = hist_data.get('历史价格列表', [])
                current_change = price_list[-1]['涨幅'] if price_list else ''
            else:
                current_change = ''
            merged[stock_code] = {'代码': stock_code, '名称': real_data.get('名称', ''), '行业': self.industry_data.get(stock_code, {}).get('行业', '-'), '现价': real_data.get('现价', ''), '涨幅': current_change, '换手率': real_data.get('换手率', ''), '流通市值': real_data.get('流通市值', ''), '今日最高价': real_data.get('今日最高价', ''), '今日最低价': real_data.get('今日最低价', ''), '涨停数_严格': limit_up_count_strict, '涨停数_宽松': limit_up_count_loose, '单日涨停数_严格': single_count_strict, '单日涨停数_宽松': single_count_loose, '连续涨停数_严格': limit_info_strict.get('最大连续涨停数', 0), '连续涨停数_宽松': limit_info_loose.get('最大连续涨停数', 0), '总涨停数_严格': max(0, total_all_limit_days_strict - limit_info_strict.get('最大连续涨停数', 0)), '总涨停数_宽松': max(0, total_all_limit_days_loose - limit_info_loose.get('最大连续涨停数', 0)), '全部涨停天数_严格': total_all_limit_days_strict, '全部涨停天数_宽松': total_all_limit_days_loose, '总涨停数_天数_严格': max(0, total_all_limit_days_strict - limit_info_strict.get('最大连续涨停数', 0)), '总涨停数_天数_宽松': max(0, total_all_limit_days_loose - limit_info_loose.get('最大连续涨停数', 0)), '离涨停多少天_严格': limit_info_strict.get('离最新日期天数', '无涨停'), '离涨停多少天_宽松': limit_info_loose.get('离最新日期天数', '无涨停'), '阳天数': sunny_days, '前N天阳天数': prev_positive_days}
            if stock_code in self.history_data:
                hist_data = self.history_data[stock_code]
                merged[stock_code].update({'昨日收盘价': hist_data.get('昨日收盘价', ''), '昨日涨幅': hist_data.get('昨日涨幅', ''), '30日最高价': hist_data.get('30日最高价', ''), '30日最低价': hist_data.get('30日最低价', ''), '60日最高价': hist_data.get('60日最高价', ''), '60日最低价': hist_data.get('60日最低价', '')})
            if stock_code in ['605188', '002337', '600262', '600403']:
                print('\n============================================================')
                print(f'{stock_code} 调试信息')
                print('============================================================')
                print(f'股票代码: {stock_code}')
                print(f"股票名称: {real_data.get('名称', '')}")
                print(f"现价: {real_data.get('现价', '')}")
                print(f"涨幅: {real_data.get('涨幅', '')}%")
                days_from_limit_strict = limit_info_strict.get('离最新日期天数', '无涨停')
                days_from_limit_loose = limit_info_loose.get('离最新日期天数', '无涨停')
                print('\n【离涨停天数】')
                print(f'  严格版: {days_from_limit_strict}')
                print(f'  宽松版: {days_from_limit_loose}')
                if days_from_limit_strict != '无涨停':
                    in_range_strict = 3 <= int(days_from_limit_strict) <= 21
                    print(f'  严格版是否在3-21范围内: {in_range_strict}')
                else:
                    print('  严格版是否在3-21范围内: False (无涨停记录)')
                if days_from_limit_loose != '无涨停':
                    in_range_loose = 3 <= int(days_from_limit_loose) <= 21
                    print(f'  宽松版是否在3-21范围内: {in_range_loose}')
                else:
                    print('  宽松版是否在3-21范围内: False (无涨停记录)')
                print(f'\n【阳天数】: {sunny_days}')
                print(f'  是否等于1: {sunny_days == 1}')
                if stock_code in self.history_data:
                    hist_data = self.history_data[stock_code]
                    yesterday_close = hist_data.get('昨日收盘价', 0)
                    realtime_price = float(real_data.get('现价', 0))
                    print(f'  昨日收盘价: {yesterday_close}')
                    print(f'  今日现价: {realtime_price}')
                    print(f"  今天vs昨天: {('上涨' if realtime_price > yesterday_close else '下跌')}")
                consecutive_limit_strict = limit_info_strict.get('最大连续涨停数', 0)
                consecutive_limit_loose = limit_info_loose.get('最大连续涨停数', 0)
                print('\n【连续涨停数】')
                print(f'  严格版（3/68开头需19.8%）: {consecutive_limit_strict}')
                print(f'  宽松版（统一9.8%）: {consecutive_limit_loose}')
                print(f'  严格版是否>=2: {consecutive_limit_strict >= 2}')
                print(f'  宽松版是否>=2: {consecutive_limit_loose >= 2}')
                if consecutive_limit_strict >= 2:
                    print(f"  严格版最后涨停日期: {limit_info_strict.get('最后涨停日期', '')}")
                if consecutive_limit_loose >= 2:
                    print(f"  宽松版最后涨停日期: {limit_info_loose.get('最后涨停日期', '')}")
                if stock_code in self.history_data:
                    hist_data = self.history_data[stock_code]
                    price_list = hist_data.get('历史价格列表', [])
                    concept_dates = set()
                    for date_str in self.concept_data.keys():
                        match = re.match('(\\d+)月(\\d+)', date_str)
                        if match:
                            month = int(match.group(1))
                            day = int(match.group(2))
                            year = 2025
                            concept_dates.add(f'{year:04d}-{month:02d}-{day:02d}')
                    print('\n【所有涨停日信息（宽松版）】')
                    limit_up_dates_loose = []
                    for i, price_data in enumerate(price_list):
                        date_str = price_data['日期']
                        if concept_dates and date_str not in concept_dates:
                            continue
                        is_zt_loose = self.is_limit_up_loose(stock_code, price_data['涨幅'])
                        if is_zt_loose:
                            limit_up_dates_loose.append({'index': i, 'date': date_str, 'change': price_data['涨幅']})
                    print(f'  涨停日总数: {len(limit_up_dates_loose)}')
                    if len(limit_up_dates_loose) > 0:
                        print('  涨停日列表:')
                        for idx, item in enumerate(limit_up_dates_loose[:10]):
                            print(f"    索引{item['index']}: 日期{item['date']}, 涨幅{item['change']}%")
                        if len(limit_up_dates_loose) > 10:
                            print(f'    ... (还有{len(limit_up_dates_loose) - 10}个)')
                    if len(limit_up_dates_loose) >= 2:
                        print('\n【连续涨停段分析】')
                        current_streak = 1
                        max_streak_found = 0
                        streak_start = 0
                        for i in range(1, len(limit_up_dates_loose)):
                            prev_index = limit_up_dates_loose[i - 1]['index']
                            curr_index = limit_up_dates_loose[i]['index']
                            if curr_index - prev_index == 1:
                                current_streak += 1
                                if current_streak > max_streak_found:
                                    max_streak_found = current_streak
                            else:
                                if current_streak >= 2:
                                    print(f'  发现连续涨停段: {current_streak}天')
                                current_streak = 1
                        if current_streak >= 2:
                            if current_streak > max_streak_found:
                                max_streak_found = current_streak
                            print(f'  最后连续涨停段: {current_streak}天')
                        print(f'  最大连续涨停数: {max_streak_found}')
                if stock_code in self.history_data:
                    hist_data = self.history_data[stock_code]
                    max_30d = hist_data.get('30日最高价', 0)
                    current_price = float(real_data.get('现价', 0))
                    if max_30d > 0:
                        percent_from_30d_high = (current_price - max_30d) / max_30d * 100
                        print('\n【突破30日新高】')
                        print(f'  30日最高价: {max_30d}')
                        print(f'  当前价格: {current_price}')
                        print(f'  离30日新高%: {percent_from_30d_high:.2f}%')
                        print(f'  是否>=100%: {percent_from_30d_high >= 100}')
                print('\n【涨停数】')
                print(f'  严格版: {limit_up_count_strict}')
                print(f'  宽松版: {limit_up_count_loose}')
                print('  (在参数范围内的涨停段数)')
                print('\n【其他筛选条件】')
                in_concept = False
                for date_str, stocks_list in self.concept_data.items():
                    for stock_data in stocks_list:
                        if stock_data[0] == stock_code:
                            in_concept = True
                            break
                    if in_concept:
                        break
                print(f'  是否在概念数据中: {in_concept}')
                print('\n【综合判断】')
                cond1_strict = days_from_limit_strict != '无涨停' and 2 <= int(days_from_limit_strict) <= 20
                cond1_loose = days_from_limit_loose != '无涨停' and 2 <= int(days_from_limit_loose) <= 20
                cond2 = sunny_days == 1
                cond3_strict = consecutive_limit_strict >= 2
                cond3_loose = consecutive_limit_loose >= 2
                cond4 = in_concept
                cond5 = True
                if stock_code in self.history_data:
                    hist_data = self.history_data[stock_code]
                    max_30d = hist_data.get('30日最高价', 0)
                    current_price = float(real_data.get('现价', 0))
                    if max_30d > 0:
                        percent_from_30d_high = (current_price - max_30d) / max_30d * 100
                        cond5 = percent_from_30d_high >= 100
                        print(f'  突破30日新高>=100%: {cond5} (实际: {percent_from_30d_high:.2f}%)')
                print(f'  离涨停3-21天（严格版）: {cond1_strict}')
                print(f'  离涨停3-21天（宽松版）: {cond1_loose}')
                print(f'  阳天数=1: {cond2}')
                print(f'  连续涨停>=2（严格版）: {cond3_strict}')
                print(f'  连续涨停>=2（宽松版）: {cond3_loose}')
                print(f'  在概念数据中: {cond4}')
                print(f'  前端全部满足（严格版）: {cond1_strict and cond2 and cond3_strict and cond4 and cond5}')
                print(f'  前端全部满足（宽松版）: {cond1_loose and cond2 and cond3_loose and cond4 and cond5}')
                print('============================================================\n')
            try:
                current_price = float(merged[stock_code].get('现价', 0))
                today_high = float(merged[stock_code].get('今日最高价', 0))
                today_low = float(merged[stock_code].get('今日最低价', 0))
                max_30d = float(merged[stock_code].get('30日最高价', 0))
                max_60d = float(merged[stock_code].get('60日最高价', 0))
                if current_price > 0 and today_high > 0:
                    merged[stock_code]['离最高价%'] = f'{(current_price - today_high) / today_high * 100:.2f}'
                else:
                    merged[stock_code]['离最高价%'] = '0.00'
                if current_price > 0 and today_low > 0:
                    merged[stock_code]['离最低价%'] = f'{(current_price - today_low) / today_low * 100:.2f}'
                else:
                    merged[stock_code]['离最低价%'] = '0.00'
                if current_price > 0 and max_30d > 0:
                    percent_to_30d = (current_price - max_30d) / max_30d * 100
                    merged[stock_code]['离30日新高%'] = f'{percent_to_30d:.2f}'
                else:
                    merged[stock_code]['离30日新高%'] = '0.00'
                if current_price > 0 and max_60d > 0:
                    percent_to_60d = (current_price - max_60d) / max_60d * 100
                    merged[stock_code]['离60日新高%'] = f'{percent_to_60d:.2f}'
                else:
                    merged[stock_code]['离60日新高%'] = '0.00'
            except:
                merged[stock_code]['离最高价%'] = '0.00'
                merged[stock_code]['离最低价%'] = '0.00'
                merged[stock_code]['离30日新高%'] = '0.00'
                merged[stock_code]['离60日新高%'] = '0.00'
        self.merged_data = merged
        return merged

    def get_merged_data(self, min_days=3, max_days=21):
        """获取合并后的数据
        
        Args:
            min_days: 离涨停天数的最小值（用于计算总涨停数）
            max_days: 离涨停天数的最大值（用于计算总涨停数）
        """
        self.merge_all_data(min_days=min_days, max_days=max_days)
        return self.merged_data

    def get_concept_count(self):
        concept_count = {}
        for date, stocks_list in self.concept_data.items():
            for stock_data in stocks_list:
                concept = stock_data[3]
                if concept and concept != '其他':
                    concept_count[concept] = concept_count.get(concept, 0) + 1
        return concept_count

    def get_today_limit_up_count(self):
        """统计每个概念的今日涨停数（使用严格标准：300/688需19.8%，其他9.8%）"""
        today_limit_up = {}
        counted_stocks = set()
        for date, stocks_list in self.concept_data.items():
            for stock_data in stocks_list:
                stock_code = stock_data[0]
                concept = stock_data[3]
                if concept and concept != '其他' and (stock_code not in counted_stocks):
                    prefix_code = None
                    if stock_code.startswith(('6', '9')):
                        prefix_code = f'sh{stock_code}'
                    elif stock_code.startswith(('0', '3', '2')):
                        prefix_code = f'sz{stock_code}'
                    if prefix_code and prefix_code in self.real_time_data:
                        real_data = self.real_time_data[prefix_code]
                        try:
                            change_pct = float(real_data.get('涨幅', 0))
                            if self.is_limit_up(stock_code, change_pct):
                                today_limit_up[concept] = today_limit_up.get(concept, 0) + 1
                                counted_stocks.add(stock_code)
                        except (ValueError, TypeError):
                            continue
        return today_limit_up

    def start_auto_update(self, interval=5):
        if self.auto_update_running:
            return {'状态': '已运行', '消息': '自动更新已在运行'}
        self.auto_update_running = True
        self.update_thread = threading.Thread(target=self._auto_update_loop, args=(interval,), daemon=True)
        self.update_thread.start()
        print('后台自动更新已启动')
        return {'状态': '已启动', '消息': '自动更新已启动'}

    def stop_auto_update(self):
        self.auto_update_running = False
        return {'状态': '已停止', '消息': '自动更新已停止'}

    def _auto_update_loop(self, interval):
        print('后台自动更新任务启动')
        while self.auto_update_running:
            time_info = get_xls_data.get_current_time_info()
            hour = time_info['小时']
            minute = time_info['分钟']
            second = time_info['秒']
            is_trading_time = False
            if hour == 9 and minute >= 15:
                is_trading_time = True
            elif 10 <= hour <= 15:
                is_trading_time = True
            elif hour == 16 and minute == 0:
                is_trading_time = True
            if hour == 9 and minute == 14 and (second >= 30):
                print(f"[{time_info['时间']}] 开始检查数据更新...")
                self._check_and_update_data()
                time.sleep(10)
            elif is_trading_time:
                print(f"[{time_info['时间']}] 交易时间更新数据...")
                self._update_all_data()
                time.sleep(interval)
            else:
                print(f"[{time_info['时间']}] 非交易时间，暂停更新...")
                time.sleep(60)

    def _check_and_update_data(self):
        try:
            if not self.real_time_data:
                self._update_all_data()
                return
            sample_codes = list(self.real_time_data.keys())[:5]
            updated = False
            for prefix_code in sample_codes:
                stock_code = prefix_code[2:]
                old_price = self.real_time_data[prefix_code].get('现价', '')
                is_updated, new_price = get_xls_data.check_data_updated(stock_code, old_price)
                if is_updated:
                    print(f'检测到数据更新: {stock_code} {old_price} -> {new_price}')
                    updated = True
                    break
            if updated:
                print('数据已更新，开始获取最新数据...')
                self._update_all_data()
            else:
                print('数据尚未更新，继续等待...')
        except Exception as e:
            print(f'检查数据更新失败: {e}')

    def _update_all_data(self):
        try:
            previewValue = int(webview.windows[0].evaluate_js('document.querySelector(".preview").value'))
            backValue = int(webview.windows[0].evaluate_js('document.querySelector(".back").value'))
            priority_codes_js = webview.windows[0].evaluate_js('window.getCurrentDisplayedStocks ? window.getCurrentDisplayedStocks() : []')
            priority_codes = priority_codes_js if priority_codes_js else []
            print(f'自动更新使用参数: preview={previewValue}, back={backValue}, 表格股票数={len(priority_codes)}')
        except Exception as e:
            print(f'获取前端参数失败，使用默认值: {e}')
            previewValue = 3
            backValue = 21
            priority_codes = []
        result = self.get_real_time_data(strat_index=None, count=backValue, show_progress=False, priority_codes=priority_codes)
        concept_count = self.get_concept_count()
        print(f'数据更新完成: {self.last_update_time} - {self.data_source_info}')
        try:
            webview.windows[0].evaluate_js('if(window.hideProgress) hideProgress();')
        except Exception as e:
            pass
        try:
            import json
            merged_data = self.get_merged_data(min_days=previewValue, max_days=backValue)
            real_time_data = self.real_time_data
            concept_data = self.concept_data
            today_limit_up = self.get_today_limit_up_count()
            merged_json = json.dumps(merged_data, ensure_ascii=False)
            real_time_json = json.dumps(real_time_data, ensure_ascii=False)
            concept_data_json = json.dumps(concept_data, ensure_ascii=False)
            concept_count_json = json.dumps(concept_count, ensure_ascii=False)
            today_limit_up_json = json.dumps(today_limit_up, ensure_ascii=False)
            webview.windows[0].evaluate_js(f'window.mergedData = {merged_json}')
            webview.windows[0].evaluate_js(f'window.realTimeData = {real_time_json}')
            webview.windows[0].evaluate_js(f'window.conceptData = {concept_data_json}')
            webview.windows[0].evaluate_js(f'window.conceptCount = {concept_count_json}')
            webview.windows[0].evaluate_js(f'window.todayLimitUp = {today_limit_up_json}')
            webview.windows[0].evaluate_js('if(window.fillStockTable) fillStockTable();')
        except Exception as e:
            print(f'推送数据到前端失败: {e}')
        except Exception as e:
            print(f'更新数据失败: {e}')
            try:
                webview.windows[0].evaluate_js('if(window.hideProgress) hideProgress();')
            except:
                pass

    def get_update_status(self):
        return {'运行中': self.auto_update_running, '最后更新': self.last_update_time, '数据源': self.data_source_info, '时间信息': get_xls_data.get_current_time_info()}
api = Api()
webview.create_window(title='股票爬虫程序', url=get_resource_path('index.html'), width=800, height=600, resizable=True, fullscreen=False, js_api=api)
webview.start(debug=True)
"""
Step 2 + 3: 处理核验结果并生成电话系统模板

- Step 2: 读取支付宝核验结果，筛选"无异常"记录，匹配公司信息
- Step 3: 生成电话系统模板Excel文件
"""

import pandas as pd
import glob
import os
import gc
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# 需排除的行业关键词（与 step1_generator 保持一致）
EXCLUDED_KEYWORDS = [
    '畜牧', '养殖',
    '水产', '渔业',
    '食品', '饲料', '屠宰', '肉类', '乳制', '饮料', '酿酒',
    '烟草', '卷烟', '香烟', '烟酒', '烟叶', '烟草制品',
    '软件',
    '种植', '农作物', '林业', '园艺',
    '餐饮', '住宿', '酒店', '居民服务', '家政', '理发', '美容', '洗浴',
    '洗染', '摄影', '旅游', '娱乐', '物业管理', '中介', '会展', '婚庆',
    '保健', '养生', '教育', '培训', '学校', '诊所', '卫生', '社会工作',
    '影视', '广电', '出版', '新闻', '演艺', '商务服务', '法律服务',
    '会计', '审计', '税务服务', '人力资源', '劳务', '广告',
    '文化艺术', '体育', '咨询服务', '快递', '邮政',
    '情趣', '成人用品', '计生用品',
    '医疗', '医药', '制药', '药业', '药材', '药品', '药物',
    '医疗器械', '医疗设备', '医疗用品', '医学',
    '医学检验', '医学影像', '诊断', '诊疗',
    '药店', '药房', '药铺', '大药房',
    '医院', '卫生院', '诊所', '门诊', '卫生室', '卫生站',
    '康复', '护理', '疗养', '体检',
    '中药', '西药', '生化药', '原料药', '生物制品',
    '保健品', '保健食品', '齿科', '口腔', '眼科', '辅助器具', '假肢',
    '农药', '化肥', '农资',
    '企业管理', '商业保理', '租赁', '拍卖', '评估', '鉴定', '认证',
    '检测', '检验', '检查', '计量', '标准化', '质量检测',
    '信用评级', '征信', '担保', '典当', '寄卖',
    '工程管理', '项目管理', '供应链管理', '商业管理',
    '物业管理', '后勤管理', '行政',
]


def is_excluded_industry(industry_str):
    """判断企业所属行业是否属于需要排除的行业"""
    if not industry_str or str(industry_str).strip() in ('', 'nan', 'None'):
        return False
    industry = str(industry_str).strip()
    for kw in EXCLUDED_KEYWORDS:
        if kw in industry:
            return True
    return False


def filter_industrial_products(df):
    """强制过滤非工业品企业，仅保留工业类企业（仅按行业筛选，不再用经营范围筛选）"""
    industry_col = None
    for col in ['所属行业', '行业', '国民经济行业', '行业分类', '行业类型']:
        if col in df.columns:
            industry_col = col
            break
    if industry_col is None:
        print("  工业品筛选: 未找到行业列，跳过过滤")
        return df, 0

    before_count = len(df)
    import pandas as pd
    mask = ~df[industry_col].apply(is_excluded_industry)
    filtered_count = before_count - mask.sum()
    df_filtered = df[mask].reset_index(drop=True)
    print(f"  工业品筛选: 过滤掉 {filtered_count} 家非工业类企业（仅按行业筛选），剩余 {len(df_filtered)} 家")
    return df_filtered, filtered_count


def _detect_check_result_header(file_path):
    """检测支付宝核验结果文件的表头行，返回最佳header索引(0-based)"""
    check_keywords = ['收款方', '账号', '姓名', '金额', '备注', '异常']
    best_header = 1  # 默认第2行(0-indexed=1)
    best_score = 0

    is_xls = file_path.lower().endswith('.xls')
    engine = 'xlrd' if is_xls else 'openpyxl'

    for try_header in [1, 0, 2, 3]:
        try:
            df_test = pd.read_excel(file_path, header=try_header, nrows=3, dtype=str, engine=engine)
            if len(df_test) == 0:
                continue
            cols = [str(c).strip() for c in df_test.columns]
            score = sum(1 for kw in check_keywords if any(kw in c for c in cols))
            if score > best_score:
                best_score = score
                best_header = try_header
            if score >= 3:
                break
        except:
            continue

    return best_header


def load_check_results(result_dir, file_pattern='*_check_result*.xls'):
    """
    加载目录中所有支付宝核验结果文件
    所有读取必须使用dtype=str防止手机号被解析为浮点数/科学计数法
    """
    # 先按指定模式找
    pattern = os.path.join(result_dir, file_pattern)
    files = sorted(glob.glob(pattern))

    if not files:
        # 也尝试一下xlsx
        pattern_xlsx = os.path.join(result_dir, '*_check_result*.xlsx')
        files = sorted(glob.glob(pattern_xlsx))

    # 如果还没找到，扫描目录下所有xls/xlsx文件
    if not files:
        all_xls = sorted(glob.glob(os.path.join(result_dir, '*.xls')))
        all_xlsx = sorted(glob.glob(os.path.join(result_dir, '*.xlsx')))
        files = all_xls + all_xlsx

    print(f"找到 {len(files)} 份核验结果文件")

    all_results = []
    for f in files:
        df = None
        basename = os.path.basename(f)

        # 自动检测表头行
        header_row = _detect_check_result_header(f)
        is_xls = f.lower().endswith('.xls')

        # 策略1: 按扩展名选择引擎 + dtype=str
        try:
            if is_xls:
                df = pd.read_excel(f, header=header_row, dtype=str, engine='xlrd')
            else:
                df = pd.read_excel(f, header=header_row, dtype=str, engine='openpyxl')
            df = df.fillna('')
            print(f"  {basename[:45]}: {len(df)} 行 (表头第{header_row+1}行)")
            all_results.append(df)
            continue
        except Exception as e:
            print(f"  读取失败(策略1): {basename[:40]} - {e}")

        # 策略2: 备用引擎 + dtype=str
        try:
            if is_xls:
                df = pd.read_excel(f, header=header_row, dtype=str, engine='openpyxl')
            else:
                df = pd.read_excel(f, header=header_row, dtype=str, engine='xlrd')
            df = df.fillna('')
            print(f"    策略2(备用引擎)成功: {len(df)} 行")
            all_results.append(df)
            continue
        except Exception as e2:
            print(f"    策略2也失败: {e2}")

        # 策略3: xlrd忽略损坏（仅xls）+ dtype=str
        if is_xls:
            try:
                book = xlrd.open_workbook(f, ignore_workbook_corruption=True)
                df = pd.read_excel(book, header=header_row, dtype=str)
                df = df.fillna('')
                print(f"    策略3(忽略损坏)成功: {len(df)} 行")
                all_results.append(df)
                continue
            except Exception as e3:
                print(f"    策略3也失败: {e3}")

        print(f"  ⚠ 无法读取文件: {basename}")

    if not all_results:
        raise Exception(f"未找到任何有效的核验结果文件: {result_dir}")

    merged = pd.concat(all_results, ignore_index=True)

    # 列名标准化
    col_map = {
        '收款方账号（必填）': '收款方账号',
        '收款方账号名称（必填）': '收款方账号名称',
        '金额（必填，单位：元）': '金额',
        '备注（选填，可输入附言/工号）': '备注',
    }
    merged = merged.rename(columns=col_map)
    merged['收款方账号'] = merged['收款方账号'].apply(_normalize_account)

    return merged


def _normalize_account(val):
    """标准化收款方账号，修复Excel读取导致的浮点数/科学计数法问题"""
    s = str(val).strip().replace('\t', '')
    if not s or s.lower() in ('nan', 'none', 'null'):
        return ''
    if s.endswith('.0') and s[:-2].isdigit():
        return s[:-2]
    if 'e+' in s.lower():
        try:
            return str(int(float(s)))
        except:
            pass
    return s


def match_company_info(no_exception_df, map_file):
    """
    将核验结果与公司信息匹配
    """
    df_map = pd.read_csv(map_file, dtype=str).fillna('')
    if '手机号' in df_map.columns:
        df_map['手机号'] = df_map['手机号'].apply(_normalize_account)
    print(f"映射记录数: {len(df_map)}")
    print(f"待匹配记录数: {len(no_exception_df)}")

    # 检查未匹配的记录
    map_phones = set(df_map['手机号'].astype(str).tolist()) if '手机号' in df_map.columns else set()
    check_phones = set(no_exception_df['收款方账号'].astype(str).tolist())
    matched_phones = check_phones & map_phones
    unmatched_phones = check_phones - map_phones
    print(f"匹配: {len(matched_phones)} 个号码, 未匹配: {len(unmatched_phones)} 个号码")

    # 合并
    result = no_exception_df.merge(df_map, left_on='收款方账号',
                                   right_on='手机号', how='left')

    # 安全检查：确保企业名称列存在
    if '企业名称' not in result.columns:
        result['企业名称'] = ''

    # 分离匹配和未匹配的记录
    matched = result[result['企业名称'].notna() & (result['企业名称'].astype(str).str.strip() != '')].copy()
    unmatched = result[result['企业名称'].isna() | (result['企业名称'].astype(str).str.strip() == '')].copy()

    # 保留未匹配记录，用手机号作为企业名称占位
    if len(unmatched) > 0:
        print(f"⚠ 未匹配记录: {len(unmatched)} 条（保留，企业名称留空）")
        unmatched['企业名称'] = unmatched['企业名称'].fillna('')

    # 按手机号+企业名称去重（避免merge导致的重复）
    df_valid = pd.concat([matched, unmatched], ignore_index=True)
    df_valid = df_valid.drop_duplicates(
        subset=['收款方账号', '企业名称'], keep='first'
    ).reset_index(drop=True)

    print(f"匹配成功: {len(matched)} 条, 未匹配保留: {len(unmatched)} 条, 合计: {len(df_valid)} 条")
    return df_valid


def filter_by_industry(df, keywords):
    """
    按行业关键词筛选（向量化加速）
    匹配"所属行业"和"经营范围"中的关键词
    """
    if not keywords:
        return df

    df = df.copy()
    industry_col = df.get('所属行业', pd.Series(dtype=str)).astype(str).fillna('')
    scope_col = df.get('经营范围', pd.Series(dtype=str)).astype(str).fillna('')
    combined = industry_col + scope_col

    matched_all = pd.Series('', index=df.index)
    for kw in keywords:
        mask = combined.str.contains(kw, na=False)
        matched_all = matched_all.where(~mask, matched_all + ('/' if matched_all.str.len() > 0 else '') + kw)

    df['匹配行业'] = matched_all
    df_filtered = df[df['匹配行业'] != ''].copy()

    print(f"行业筛选: {len(df)} -> {len(df_filtered)} 条")
    print(f"  关键词: {', '.join(keywords)}")

    return df_filtered


def merge_phones_by_company(df):
    """
    按企业名称分组，多个号码合并到一行用/连接
    注意：企业名称为空的记录不合并，逐条保留（避免未匹配记录被合并丢失）
    """
    # 安全检查：确保企业名称列存在
    if '企业名称' not in df.columns:
        df['企业名称'] = ''

    # 分离有企业名称和无企业名称的记录
    has_name = df[df['企业名称'].astype(str).str.strip() != ''].copy()
    no_name = df[df['企业名称'].astype(str).str.strip() == ''].copy()

    grouped_data = []

    # 有名称的按公司合并
    if len(has_name) > 0:
        for company, group in has_name.groupby('企业名称', sort=False):
            phones = '/'.join(sorted(set(group['收款方账号'].tolist())))
            first = group.iloc[0]
            grouped_data.append({
                '企业名称': company,
                '法定代表人': first.get('法定代表人', ''),
                '收款方账号': phones,
                '注册资本': first.get('注册资本', ''),
                '实缴资本': first.get('实缴资本', ''),
                '成立日期': first.get('成立日期', ''),
                '经营状态': first.get('经营状态', ''),
                '所属行业': first.get('所属行业', ''),
                '注册地址': first.get('注册地址', ''),
                '经营范围': first.get('经营范围', ''),
                '统一社会信用代码': first.get('统一社会信用代码', ''),
                '参保人数': first.get('参保人数', ''),
                '所属省份': first.get('所属省份', ''),
                '所属城市': first.get('所属城市', ''),
                '所属区县': first.get('所属区县', ''),
                '官网': first.get('官网', '') or first.get('企业官网', '') or first.get('网站', '') or first.get('网址', ''),
                '匹配行业': first.get('匹配行业', ''),
                '股份信息': first.get('股份信息', '') or first.get('股东信息', '') or first.get('大股东', '') or '',
            })

    # 无名称的逐条保留（用手机号作为标识）
    if len(no_name) > 0:
        for _, row in no_name.iterrows():
            grouped_data.append({
                '企业名称': '',
                '法定代表人': row.get('法定代表人', ''),
                '收款方账号': row.get('收款方账号', ''),
                '注册资本': row.get('注册资本', ''),
                '实缴资本': row.get('实缴资本', ''),
                '成立日期': row.get('成立日期', ''),
                '经营状态': row.get('经营状态', ''),
                '所属行业': row.get('所属行业', ''),
                '注册地址': row.get('注册地址', ''),
                '经营范围': row.get('经营范围', ''),
                '统一社会信用代码': row.get('统一社会信用代码', ''),
                '参保人数': row.get('参保人数', ''),
                '所属省份': row.get('所属省份', ''),
                '所属城市': row.get('所属城市', ''),
                '所属区县': row.get('所属区县', ''),
                '官网': row.get('官网', '') or row.get('企业官网', '') or row.get('网站', '') or row.get('网址', ''),
                '匹配行业': row.get('匹配行业', ''),
                '股份信息': row.get('股份信息', '') or row.get('股东信息', '') or row.get('大股东', '') or '',
            })

    result_df = pd.DataFrame(grouped_data)
    print(f"按公司合并: {len(df)} 条 -> {len(result_df)} 条 (有名称合并: {len(grouped_data) - len(no_name)} 家, 无名称保留: {len(no_name)} 条)")
    return result_df


def create_phone_template(df, output_path, remark='', operator_name='陈平安'):
    """
    生成电话系统模板Excel
    """
    template_headers = [
        '公司名称', '法人姓名', '法人电话', '注册资本', '实缴资本', '成立日期',
        '经营状态', '所属行业', '注册地址', '经营范围', '统一社会信用代码',
        '公司类型', '参保人数', '所属省份', '所属城市', '所属区县',
        '官网', '股份信息', '资料人员', '号码来源', '备注'
    ]

    # 构建数据
    template_data = []
    for _, r in df.iterrows():
        # 确定备注内容
        remark_val = remark
        if not remark_val and '匹配行业' in df.columns:
            remark_val = str(r.get('匹配行业', ''))

        template_data.append({
            '公司名称': r.get('企业名称', ''),
            '法人姓名': r.get('法定代表人', ''),
            '法人电话': r.get('收款方账号', ''),
            '注册资本': r.get('注册资本', ''),
            '实缴资本': r.get('实缴资本', ''),
            '成立日期': r.get('成立日期', ''),
            '经营状态': r.get('经营状态', ''),
            '所属行业': r.get('所属行业', ''),
            '注册地址': r.get('注册地址', ''),
            '经营范围': r.get('经营范围', ''),
            '统一社会信用代码': r.get('统一社会信用代码', ''),
            '公司类型': '',
            '参保人数': r.get('参保人数', ''),
            '所属省份': r.get('所属省份', ''),
            '所属城市': r.get('所属城市', ''),
            '所属区县': r.get('所属区县', ''),
            '官网': r.get('官网', '') or r.get('企业官网', '') or r.get('网站', '') or r.get('网址', ''),
            '股份信息': r.get('股份信息', '') or r.get('股东信息', '') or r.get('大股东', '') or '',
            '资料人员': operator_name,
            '号码来源': '',
            '备注': remark_val,
        })

    # 创建Excel
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    # 样式
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    thin_border = Border(
        left=Side(style='thin', color='D9DEE7'),
        right=Side(style='thin', color='D9DEE7'),
        top=Side(style='thin', color='D9DEE7'),
        bottom=Side(style='thin', color='D9DEE7')
    )
    zebra_fill = PatternFill(start_color='F7F9FC', end_color='F7F9FC', fill_type='solid')
    center_align = Alignment(horizontal='center', vertical='center')
    data_font = Font(name='微软雅黑', size=10)

    # 表头
    for c, h in enumerate(template_headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # 数据 - 批量写入值，再统一设置样式
    for ri, row in enumerate(template_data):
        excel_row = ri + 2
        for ci, h in enumerate(template_headers):
            ws.cell(excel_row, ci + 1, str(row.get(h, '')))

    # 批量设置数据区样式
    max_row = len(template_data) + 1
    max_col = len(template_headers)
    for ri in range(2, max_row + 1):
        is_zebra = (ri - 2) % 2 == 1
        for ci in range(1, max_col + 1):
            cell = ws.cell(ri, ci)
            cell.border = thin_border
            cell.font = data_font
            if is_zebra:
                cell.fill = zebra_fill

    # 列宽
    col_widths = {
        '公司名称': 35, '法人姓名': 12, '法人电话': 25, '注册资本': 12,
        '实缴资本': 12, '成立日期': 12, '经营状态': 10, '所属行业': 18,
        '注册地址': 40, '经营范围': 50, '统一社会信用代码': 22, '公司类型': 20,
        '参保人数': 10, '所属省份': 10, '所属城市': 10, '所属区县': 10,
        '官网': 30, '股份信息': 10, '资料人员': 10, '号码来源': 10, '备注': 18
    }
    for c, h in enumerate(template_headers, 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = col_widths.get(h, 12)

    ws.freeze_panes = 'A2'
    wb.save(output_path)
    print(f"模板已保存: {os.path.basename(output_path)}")
    print(f"  共 {len(template_data)} 条记录")

    return template_data


def create_failed_list(failed_df, map_file, output_path):
    """
    生成核验失败记录列表（带公司信息）
    """
    try:
        df_map = pd.read_csv(map_file, dtype=str).fillna('')
        if '手机号' in df_map.columns:
            df_map['手机号'] = df_map['手机号'].apply(_normalize_account)
        result = failed_df.merge(df_map, left_on='收款方账号',
                                 right_on='手机号', how='left')
        result = result.drop_duplicates(
            subset=['收款方账号', '企业名称'], keep='first'
        ).reset_index(drop=True)
    except:
        result = failed_df.copy()

    out_cols = ['收款方账号', '收款方账号名称', '金额', '账户异常原因',
                '企业名称', '法定代表人', '注册资本', '成立日期',
                '所属行业', '注册地址', '统一社会信用代码', '经营范围',
                '所属省份', '所属城市', '所属区县', '参保人数']
    available = [c for c in out_cols if c in result.columns]
    df_out = result[available].reset_index(drop=True)

    wb = Workbook()
    ws = wb.active
    ws.title = '匹配失败'

    header_fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    thin_border = Border(
        left=Side(style='thin', color='D9DEE7'),
        right=Side(style='thin', color='D9DEE7'),
        top=Side(style='thin', color='D9DEE7'),
        bottom=Side(style='thin', color='D9DEE7')
    )
    zebra_fill = PatternFill(start_color='FFF2F2', end_color='FFF2F2', fill_type='solid')

    for c, h in enumerate(available, 1):
        cell = ws.cell(1, c, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    for ri, row in df_out.iterrows():
        for ci, h in enumerate(available):
            cell = ws.cell(ri + 2, ci + 1, str(row.get(h, '')))
            cell.border = thin_border
            cell.font = Font(name='微软雅黑', size=10)
            if ri % 2 == 1:
                cell.fill = zebra_fill

    col_widths = {
        '收款方账号': 15, '收款方账号名称': 12, '金额': 8, '账户异常原因': 30,
        '企业名称': 35, '法定代表人': 10, '注册资本': 12, '成立日期': 12,
        '所属行业': 15, '注册地址': 40, '统一社会信用代码': 22, '经营范围': 50,
        '所属省份': 10, '所属城市': 10, '所属区县': 10, '参保人数': 10
    }
    for c, h in enumerate(available, 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = col_widths.get(h, 12)

    ws.freeze_panes = 'A2'
    wb.save(output_path)
    print(f"失败记录已保存: {os.path.basename(output_path)}")
    print(f"  共 {len(df_out)} 条失败记录")

    return df_out


def filter_by_shareholder(df, shareholder_filter):
    """
    按大股东信息筛选
    shareholder_filter: 'yes'=仅保留有大股东, 'no'=仅保留无大股东, 其他=不筛选
    """
    if not shareholder_filter or shareholder_filter not in ('yes', 'no'):
        return df

    # 支持多种列名
    shareholder_col = None
    for col in ['股份信息', '股东信息', '大股东', '股东名称', '主要股东', '股东']:
        if col in df.columns:
            shareholder_col = col
            break

    if shareholder_col is None:
        print(f"  数据中无股东信息列，跳过大股东筛选")
        return df

    before = len(df)
    if shareholder_filter == 'yes':
        df_filtered = df[df[shareholder_col].astype(str).str.strip() != ''].copy()
        print(f"  大股东筛选(仅保留有): {before} -> {len(df_filtered)} 条")
    else:  # 'no'
        df_filtered = df[df[shareholder_col].astype(str).str.strip() == ''].copy()
        print(f"  大股东筛选(仅保留无): {before} -> {len(df_filtered)} 条")

    return df_filtered


def filter_by_insured(df, insured_filter):
    """
    按参保人数筛选
    insured_filter: 用户输入的最低参保人数（字符串），保留参保人数 >= 该值的企业
    留空或不为数字则不筛选
    """
    if not insured_filter:
        return df

    try:
        threshold = int(float(str(insured_filter).strip()))
    except (ValueError, TypeError):
        print(f"  参保人数筛选值无效('{insured_filter}')，跳过")
        return df

    # 支持多种列名
    insured_col = None
    for col in ['参保人数', '社保人数', '员工人数', '人员规模']:
        if col in df.columns:
            insured_col = col
            break

    if insured_col is None:
        print(f"  数据中无参保人数列，跳过参保人数筛选")
        return df

    def parse_number(val):
        """解析数值，支持'10人'、'10-20人'、'50人以上'等格式"""
        s = str(val).strip()
        if not s or s.lower() in ('nan', 'none', 'null', ''):
            return 0
        # 去掉常见后缀
        s = s.replace('人', '').replace('以上', '').replace('以下', '').strip()
        # 尝试取范围的最小值，如 "10-20" -> 10
        if '-' in s:
            parts = s.split('-')
            try:
                return int(float(parts[0].strip()))
            except:
                pass
        # 直接转数字
        try:
            return int(float(s))
        except:
            return 0

    before = len(df)
    df = df.copy()
    df['_insured_num'] = df[insured_col].apply(parse_number)
    df_filtered = df[df['_insured_num'] >= threshold].copy()
    df_filtered = df_filtered.drop(columns=['_insured_num'])
    print(f"  参保人数筛选(≥{threshold}): {before} -> {len(df_filtered)} 条")

    return df_filtered


def process_results(result_dir, map_file, output_prefix,
                    merge_phones=False, industry_keywords=None,
                    remark='', operator_name='陈平安', output_dir=None,
                    shareholder_filter='', insured_filter='',
                    filter_industrial=False):
    """
    完整的 Step 2 + 3 处理流程

    参数:
        result_dir: 核验结果文件目录
        map_file: 手机号映射文件（csv）
        output_prefix: 输出文件前缀
        merge_phones: 是否按公司合并多号码
        industry_keywords: 行业筛选关键词列表
        remark: 备注内容
        operator_name: 资料人员姓名
        output_dir: 输出目录
        shareholder_filter: 大股东筛选
        insured_filter: 参保人数筛选
        filter_industrial: 是否强制过滤非工业品企业

    返回:
        dict: 统计信息
    """
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(map_file))
        if not output_dir:
            output_dir = '.'

    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, f'电话系统模板_{output_prefix}_已填写.xlsx')
    failed_file = os.path.join(output_dir, f'匹配失败_{output_prefix}.xlsx')

    # Step 2: 加载核验结果
    print("\n=== Step 2: 处理核验结果 ===")
    merged = load_check_results(result_dir)
    print(f"合计: {len(merged)} 条记录")

    # 安全检查：确保账户异常原因列存在
    if '账户异常原因' not in merged.columns:
        print("  警告: 核验结果文件中未找到'账户异常原因'列，默认全部视为无异常")
        merged['账户异常原因'] = '无异常'

    # 统计
    print("\n账户异常原因分布:")
    reason_counts = merged['账户异常原因'].value_counts()
    for reason, count in reason_counts.items():
        print(f"  {reason}: {count} 条")

    # 筛选无异常
    no_exception = merged[
        merged['账户异常原因'].astype(str).str.strip() == '无异常'
    ].copy()
    print(f"\n无异常记录: {len(no_exception)} 条")
    exception_count = len(merged) - len(no_exception)
    if exception_count > 0:
        print(f"异常记录(过滤掉): {exception_count} 条")

    # 匹配公司信息
    print("\n正在匹配公司信息...")
    df_matched = match_company_info(no_exception, map_file)

    # 强制工业品筛选
    industrial_filtered = 0
    if filter_industrial:
        print("\n正在强制工业品筛选...")
        df_matched, industrial_filtered = filter_industrial_products(df_matched)

    # 行业筛选
    if industry_keywords:
        print("\n正在按行业筛选...")
        df_matched = filter_by_industry(df_matched, industry_keywords)

    # 大股东筛选
    if shareholder_filter:
        print("\n正在按大股东筛选...")
        df_matched = filter_by_shareholder(df_matched, shareholder_filter)

    # 参保人数筛选
    if insured_filter:
        print("\n正在按参保人数筛选...")
        df_matched = filter_by_insured(df_matched, insured_filter)

    # 按公司合并号码
    df_final = df_matched
    if merge_phones:
        print("\n正在按公司合并多号码...")
        print(f"  合并前: {len(df_matched)} 条")
        df_final = merge_phones_by_company(df_matched)
        print(f"  合并后: {len(df_final)} 条 (减少 {len(df_matched) - len(df_final)} 条)")
    else:
        print(f"\n最终记录数: {len(df_final)} 条")

    # Step 3: 生成模板
    print("\n=== Step 3: 生成电话系统模板 ===")
    create_phone_template(df_final, output_file, remark=remark,
                          operator_name=operator_name)

    # 生成失败记录
    failed = merged[
        merged['账户异常原因'].astype(str).str.strip() != '无异常'
    ].copy()
    if len(failed) > 0:
        print("\n正在生成失败记录...")
        create_failed_list(failed, map_file, failed_file)
    else:
        failed_file = None

    final_count = len(df_final)
    total_verified_count = len(merged)
    no_exception_count = len(no_exception)
    matched_count = len(df_matched)

    # 释放中间数据内存
    del merged, no_exception, df_matched, df_final
    if 'failed' in dir():
        del failed
    gc.collect()

    return {
        'total_verified': total_verified_count,
        'no_exception': no_exception_count,
        'final_count': final_count,
        'output_file': output_file,
        'failed_file': failed_file,
        'industrial_filtered': industrial_filtered,
        'matched_count': matched_count,
    }


if __name__ == '__main__':
    import sys
    if len(sys.argv) >= 4:
        result = process_results(
            result_dir=sys.argv[1],
            map_file=sys.argv[2],
            output_prefix=sys.argv[3],
            merge_phones='--merge' in sys.argv,
        )
        print(f"\n完成! 共 {result['final_count']} 条记录")

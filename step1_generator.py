"""
Step 1: 生成支付宝核验文件

从企业数据Excel文件中提取手机号，展开为每行一个手机号，
生成支付宝批量核验所需的Excel文件。
"""

import pandas as pd
import random
import re
import os
import gc
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# 需排除的行业关键词（畜牧业、养殖业、水产业、食品业、软件、种植业、服务业、医疗、农药）
# 匹配这些关键词的企业将被过滤掉，仅保留工业品企业
EXCLUDED_KEYWORDS = [
    # 畜牧业、养殖业
    '畜牧', '养殖',
    # 水产业
    '水产', '渔业',
    # 食品业（含农副食品、饲料、屠宰、肉类、乳制品、饮料、酒、烟草）
    '食品', '饲料', '屠宰', '肉类', '乳制', '饮料', '酿酒',
    '烟草', '卷烟', '香烟', '烟酒', '烟叶', '烟草制品',
    # 软件业
    '软件',
    # 种植业、林业
    '种植', '农作物', '林业', '园艺',
    # 服务业
    '餐饮', '住宿', '酒店', '居民服务', '家政', '理发', '美容', '洗浴',
    '洗染', '摄影', '旅游', '娱乐', '物业管理', '中介', '会展', '婚庆',
    '保健', '养生', '教育', '培训', '学校', '诊所', '卫生', '社会工作',
    '影视', '广电', '出版', '新闻', '演艺', '商务服务', '法律服务',
    '会计', '审计', '税务服务', '人力资源', '劳务', '广告',
    '文化艺术', '体育', '咨询服务', '快递', '邮政',
    # 情趣用品
    '情趣', '成人用品', '计生用品',
    # 医疗相关行业（含医药、制药、医疗器械、医疗、医院、诊所、保健等）
    '医疗', '医药', '制药', '药业', '药材', '药品', '药物',
    '医疗器械', '医疗设备', '医疗用品', '医学',
    '医学检验', '医学影像', '诊断', '诊疗',
    '药店', '药房', '药铺', '大药房',
    '医院', '卫生院', '诊所', '门诊', '卫生室', '卫生站',
    '康复', '护理', '疗养', '体检',
    '中药', '西药', '生化药', '原料药', '生物制品',
    '保健品', '保健食品', '齿科', '口腔', '眼科', '辅助器具', '假肢',
    # 农药行业
    '农药', '化肥', '农资',
    # 企业管理、检查服务相关
    '企业管理', '商业保理', '租赁', '拍卖', '评估', '鉴定', '认证',
    '检测', '检验', '检查', '计量', '标准化', '质量检测',
    '信用评级', '征信', '担保', '典当', '寄卖',
    '工程管理', '项目管理', '供应链管理', '商业管理',
    '物业管理', '后勤管理', '行政',
]


def is_excluded_industry(industry_str):
    """判断企业所属行业是否属于需要排除的行业（畜牧/养殖/水产/食品/软件/种植/服务/医疗/农药）"""
    if not industry_str or str(industry_str).strip() in ('', 'nan', 'None'):
        return False
    industry = str(industry_str).strip()
    for kw in EXCLUDED_KEYWORDS:
        if kw in industry:
            return True
    return False

def is_excluded_by_scope(scope_str):
    """判断企业经营范围是否包含需要排除的业务"""
    if not scope_str or str(scope_str).strip() in ('', 'nan', 'None'):
        return False
    scope = str(scope_str).strip()
    for kw in EXCLUDED_KEYWORDS:
        if kw in scope:
            return True
    return False


def detect_header_and_read(file_path, header_row_hint=2):
    """
    自动检测表头行并读取Excel文件
    支持爱企查、天眼查等不同导出格式
    返回: (DataFrame, 实际使用的header_row)
    """
    # 列名映射：把各种别名统一成标准列名
    col_rename_map = {
        '系统匹配企业名称': '企业名称',
        '公司名称': '企业名称',
        '名称': '企业名称',
        '企业名称.1': '企业名称',
        '企业': '企业名称',
        '公司': '企业名称',
        '单位名称': '企业名称',
        '机构名称': '企业名称',
        '商户名称': '企业名称',
        '店铺名称': '企业名称',
        '法定代表人.1': '法定代表人',
        '法人': '法定代表人',
        '负责人': '法定代表人',
        '负责人姓名': '法定代表人',
        '法人代表': '法定代表人',
        '手机号': '手机',
        '手机号码': '手机',
        '联系电话': '电话',
        '法人手机': '法人电话',
        '最终受益人名称': '股东信息',
        '受益人': '股东信息',
        '大股东': '股东信息',
    }
    
    def score_header(df):
        """给表头行打分，分数越高越可能是真正的表头"""
        score = 0
        cols = [str(c).strip() for c in df.columns]
        # 关键列名得分
        good_cols = ['企业名称', '公司名称', '系统匹配企业名称', '法定代表人', '电话', '手机',
                     '最终受益人名称', '持股比例', '成立日期', '注册资本', '统一社会信用代码']
        for c in cols:
            for kw in good_cols:
                if c == kw or c.startswith(kw):
                    score += 2
            if c.startswith('Unnamed'):
                score -= 1
        # 数据行质量：非空值比例
        if len(df) > 0:
            first_row = df.iloc[0].astype(str)
            non_empty = (first_row != '') & (~first_row.str.startswith('Unnamed')) & (~first_row.str.contains('声明|数据导出|仅供参考'))
            score += non_empty.sum()
        return score
    
    best_df = None
    best_score = -1
    best_header = header_row_hint
    
    # 尝试多行作为表头，选择得分最高的
    for try_header in [header_row_hint] + [1, 2, 0, 3, 4]:
        if try_header < 0:
            continue
        try:
            df = pd.read_excel(file_path, header=try_header, dtype=str, engine='openpyxl').fillna('')
            if len(df) == 0:
                continue
            s = score_header(df)
            if s > best_score:
                best_score = s
                best_df = df
                best_header = try_header
        except:
            pass
    
    if best_df is None:
        # 所有尝试都失败，用提示的行读取
        best_df = pd.read_excel(file_path, header=header_row_hint, dtype=str, engine='openpyxl').fillna('')
        best_header = header_row_hint
    
    # 重命名列
    rename_dict = {}
    for old_name in best_df.columns:
        old_str = str(old_name).strip()
        if old_str in col_rename_map and col_rename_map[old_str] not in [str(c).strip() for c in best_df.columns]:
            rename_dict[old_name] = col_rename_map[old_str]
    if rename_dict:
        best_df = best_df.rename(columns=rename_dict)

    # 智能回退：如果仍没有'企业名称'列，尝试模糊匹配
    if '企业名称' not in best_df.columns:
        company_keywords = ['企业名称', '公司名称', '企业', '公司', '名称', '单位名称', '机构', '商户', '店铺']
        for col in best_df.columns:
            col_str = str(col).strip()
            # 跳过Unnamed列
            if col_str.startswith('Unnamed'):
                continue
            for kw in company_keywords:
                if kw in col_str:
                    best_df = best_df.rename(columns={col: '企业名称'})
                    print(f"  智能匹配: '{col_str}' -> '企业名称'")
                    break
            if '企业名称' in best_df.columns:
                break

    # 智能回退：如果仍没有'法定代表人'列，尝试模糊匹配
    if '法定代表人' not in best_df.columns:
        legal_keywords = ['法定代表人', '法人', '负责人', '代表人']
        for col in best_df.columns:
            col_str = str(col).strip()
            if col_str.startswith('Unnamed'):
                continue
            for kw in legal_keywords:
                if kw in col_str:
                    best_df = best_df.rename(columns={col: '法定代表人'})
                    break
            if '法定代表人' in best_df.columns:
                break

    print(f"  [读取] openpyxl 成功，自动检测表头在第 {best_header + 1} 行 (得分:{best_score})")
    return best_df, best_header


def read_excel_auto(file_path, header_row=2):
    """
    自动识别文件格式并读取，支持xlsx和损坏的xls
    header_row: 表头所在行索引（0-based），默认第3行即索引2
    已废弃，保留兼容，内部调用detect_header_and_read
    """
    df, _ = detect_header_and_read(file_path, header_row)
    return df


def extract_phones(row):
    """
    从一行数据中提取所有手机号
    支持多种列名: 电话, 更多电话, 法人电话, 手机号码, 联系电话, 手机等
    列名采用模糊匹配（包含关键词即可）
    """
    phones = []

    # 电话列关键词（模糊匹配，包含这些关键词的列都会被扫描）
    phone_keywords = ['手机号', '手机号码', '联系手机', '更多电话', '联系电话', '推荐电话',
                      '法人电话', '电话1', '电话2', '有效手机号', '手机', '电话']

    for col in row.index:
        col_str = str(col)
        is_phone_col = False
        for kw in phone_keywords:
            if kw in col_str:
                is_phone_col = True
                break
        if not is_phone_col:
            continue

        phone_str = str(row[col]).strip()
        if phone_str and phone_str != 'nan':
            for p in re.findall(r'1\d{10}', phone_str):
                phones.append(p)

    # 去重并保持顺序
    seen = set()
    unique = []
    for p in phones:
        if p not in seen:
            seen.add(p)
            unique.append(p)

    return unique


def generate_alipay_files(data_file, output_prefix, header_row=3,
                          max_rows_per_file=3000, random_seed=100,
                          output_dir=None, shareholder_file=None,
                          verified_phones=None, filter_industrial=False,
                          verified_companies=None, db_path=None):
    """
    生成支付宝核验文件

    参数:
        data_file: 输入的企业数据Excel文件路径
        output_prefix: 输出文件前缀名称
        header_row: 表头所在行（1-based，默认第3行）
        max_rows_per_file: 每份文件最大行数（默认3000）
        random_seed: 随机金额种子
        output_dir: 输出目录，默认为data_file所在目录
        shareholder_file: 股东信息文件路径（可选），按公司名称匹配合并
        verified_phones: set，已核验通过的电话号码集合（内存模式，兼容旧版）
        filter_industrial: bool，是否强制过滤工业品企业
        verified_companies: set，已核验通过的企业名称集合（内存模式，兼容旧版）
        db_path: str，主数据库路径（数据库模式，优先使用，不占内存）

    返回:
        dict: 包含统计信息的字典
    """
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(data_file))
        if not output_dir:
            output_dir = '.'

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 输出文件基础路径
    out_base = os.path.join(output_dir, f'支付宝核验_{output_prefix}_已填写')
    map_file = os.path.join(output_dir, f'_phone_map_{output_prefix}.csv')

    # 读取数据（自动检测表头）
    print(f"正在读取数据文件...")
    df, actual_header = detect_header_and_read(data_file, header_row - 1)
    print(f"  原始数据: {len(df)} 行")

    # 过滤掉"进入企业店铺"和空企业名
    if '企业名称' not in df.columns:
        raise ValueError("上传的Excel文件中未找到'企业名称'列（也无法自动匹配类似列名）。请检查文件格式，确保包含企业名称/公司名称/名称等列。")
    df['企业名称'] = df['企业名称'].astype(str).str.strip()
    df = df[~df['企业名称'].astype(str).str.contains('进入企业店铺', na=False)]
    df = df[df['企业名称'] != '']
    df = df[~df['企业名称'].str.startswith('Unnamed')]
    print(f"  过滤后: {len(df)} 行")

    # 强制过滤非工业品企业（畜牧/养殖/水产/食品/软件/种植/医疗/农药/烟草等），仅保留工业品
    industrial_filtered = 0
    if filter_industrial:
        industry_col = None
        for col in ['所属行业', '行业', '国民经济行业', '行业分类', '行业类型']:
            if col in df.columns:
                industry_col = col
                break

        if industry_col:
            before_count = len(df)
            mask = ~df[industry_col].apply(is_excluded_industry)
            industrial_filtered = before_count - mask.sum()
            df = df[mask].reset_index(drop=True)
            print(f"  工业品筛选: 过滤掉 {industrial_filtered} 家非工业类企业（仅按行业筛选），剩余 {len(df)} 家")
        else:
            print("  工业品筛选: 未找到行业列，跳过过滤")

    # 过滤掉资料库中已核验通过的企业（按公司名称精确匹配）
    company_filtered = 0
    if db_path:
        # 数据库模式：用SQLite临时表过滤，不占内存
        from db_filter import filter_verified_companies
        company_names = df['企业名称'].astype(str).str.strip().tolist()
        verified_set, company_filtered = filter_verified_companies(db_path, company_names)
        if verified_set:
            before_count = len(df)
            mask = ~df['企业名称'].astype(str).str.strip().isin(verified_set)
            df = df[mask].reset_index(drop=True)
            print(f"  资料库匹配: 过滤掉 {company_filtered} 家已核验通过的企业，剩余 {len(df)} 家 (数据库模式)")
            del verified_set
            gc.collect()
        else:
            print("  资料库匹配: 无已核验企业数据，跳过 (数据库模式)")
        del company_names
        gc.collect()
    elif verified_companies:
        # 内存模式（兼容旧版）
        before_count = len(df)
        df['企业名称_lower'] = df['企业名称'].astype(str).str.strip()
        mask = ~df['企业名称_lower'].isin(verified_companies)
        company_filtered = before_count - mask.sum()
        df = df[mask].reset_index(drop=True)
        del df['企业名称_lower']
        print(f"  资料库匹配: 过滤掉 {company_filtered} 家已核验通过的企业，剩余 {len(df)} 家")
    else:
        print("  资料库匹配: 无已核验企业数据，跳过")

    # 提取并展开手机号（向量化，比iterrows快10倍以上）
    print(f"\n正在提取手机号...")
    # 电话列关键词（模糊匹配，包含这些关键词的列都会被扫描）
    phone_keywords = ['手机号', '手机号码', '联系手机', '更多电话', '联系电话', '推荐电话',
                      '法人电话', '电话1', '电话2', '有效手机号', '手机', '电话']
    available_phone_cols = [c for c in df.columns if any(kw in str(c) for kw in phone_keywords)]
    
    if available_phone_cols:
        # 合并所有电话列文本，用正则一次性提取所有手机号
        def extract_all_phones(row_vals):
            text = ' '.join(str(v).strip() for v in row_vals if str(v).strip() and str(v).strip() != 'nan')
            phones = re.findall(r'1\d{10}', text)
            # 去重保序
            seen = set()
            unique = []
            for p in phones:
                if p not in seen:
                    seen.add(p)
                    unique.append(p)
            return unique
        
        df['_phones'] = df[available_phone_cols].apply(extract_all_phones, axis=1)
        df_expanded = df.explode('_phones').rename(columns={'_phones': '手机号'}).reset_index(drop=True)
        df_expanded = df_expanded[df_expanded['手机号'].notna() & (df_expanded['手机号'] != '')]
    else:
        # 没有电话列，保持原样，留空手机号
        df_expanded = df.copy()
        df_expanded['手机号'] = ''
    
    print(f"  展开后（每个手机号一行）: {len(df_expanded)} 条")

    # 过滤掉已核验通过的电话号码（跳过不需要重新核验的）
    skipped_count = 0
    if db_path:
        # 数据库模式：用SQLite临时表过滤，不占内存
        from db_filter import filter_verified_phones_fast
        phone_list = df_expanded['手机号'].tolist()
        verified_phones_set, skipped_count = filter_verified_phones_fast(db_path, phone_list)
        if verified_phones_set:
            mask = ~df_expanded['手机号'].isin(verified_phones_set)
            df_expanded = df_expanded[mask].reset_index(drop=True)
            print(f"  已核验通过跳过: {skipped_count} 条 (数据库模式)")
            print(f"  需要核验: {len(df_expanded)} 条")
            del verified_phones_set
            gc.collect()
        del phone_list
        gc.collect()
    elif verified_phones:
        before_count = len(df_expanded)
        mask = ~df_expanded['手机号'].isin(verified_phones)
        skipped_count = before_count - mask.sum()
        df_expanded = df_expanded[mask].reset_index(drop=True)
        print(f"  已核验通过跳过: {skipped_count} 条")
        print(f"  需要核验: {len(df_expanded)} 条")

    # 保存映射文件（用于后续匹配）
    # 先统一列名
    std_rename = {
        '登记状态': '经营状态',
        '注册地址': '注册地址',
        '地址': '注册地址',
        '行业': '所属行业',
        '省': '所属省份',
        '市': '所属城市',
        '区县': '所属区县',
        '社保人数': '参保人数',
        '员工人数': '参保人数',
    }
    for old_name, new_name in std_rename.items():
        if old_name in df_expanded.columns and new_name not in df_expanded.columns:
            df_expanded = df_expanded.rename(columns={old_name: new_name})
    
    map_cols = ['企业名称', '法定代表人', '手机号', '注册资本', '实缴资本',
                '成立日期', '所属行业', '注册地址', '统一社会信用代码',
                '经营范围', '所属省份', '所属城市', '所属区县', '参保人数', '经营状态',
                '官网', '企业官网', '网站', '网址',
                '股份信息', '股东信息', '大股东', '股东名称', '主要股东']
    available_cols = [c for c in map_cols if c in df_expanded.columns]

    # 如果有股东信息文件，按公司名称匹配合并
    if shareholder_file:
        try:
            print(f"\n正在读取股东信息文件: {os.path.basename(shareholder_file)}")
            df_sh, _ = detect_header_and_read(shareholder_file, 0)
            print(f"  股东信息: {len(df_sh)} 行")

            # 查找公司名称列
            sh_company_col = None
            for col in ['企业名称', '公司名称', '名称']:
                if col in df_sh.columns:
                    sh_company_col = col
                    break

            if sh_company_col is None:
                print("  警告: 股东信息文件中未找到公司名称列，跳过合并")
            else:
                # 统一公司名称列名为"企业名称"
                if sh_company_col != '企业名称':
                    df_sh = df_sh.rename(columns={sh_company_col: '企业名称'})

                # 清理企业名称列（向下填充空值，处理多行股权路径格式）
                df_sh['企业名称'] = df_sh['企业名称'].astype(str).str.strip()
                df_sh['企业名称'] = df_sh['企业名称'].replace(['', 'nan', 'None'], None)
                df_sh['企业名称'] = df_sh['企业名称'].ffill()

                # 查找股东信息列（支持最终受益人格式）
                sh_info_col = None
                sh_ratio_col = None
                for col in ['股份信息', '股东信息', '大股东', '股东名称', '主要股东', '股东']:
                    if col in df_sh.columns:
                        sh_info_col = col
                        break
                # 爱企查受益人格式
                if '最终受益人名称' in df_sh.columns:
                    sh_info_col = '最终受益人名称'
                if '持股比例' in df_sh.columns:
                    sh_ratio_col = '持股比例'

                if sh_info_col is None:
                    print("  警告: 股东信息文件中未找到股东相关列，跳过合并")
                else:
                    # 按企业名称分组，合并股东信息
                    sh_grouped = []
                    for company, group in df_sh.groupby('企业名称', sort=False):
                        if not company or company.lower() in ('nan', 'none', 'unnamed: 0'):
                            continue
                        shareholders = []
                        for _, r in group.iterrows():
                            name = str(r.get(sh_info_col, '')).strip()
                            ratio = str(r.get(sh_ratio_col, '')).strip() if sh_ratio_col else ''
                            if name and name.lower() not in ('nan', 'none', ''):
                                if ratio and ratio.lower() not in ('nan', 'none', '-', ''):
                                    shareholders.append(f"{name}({ratio})")
                                else:
                                    shareholders.append(name)
                        sh_str = '; '.join(shareholders) if shareholders else ''
                        sh_grouped.append({'企业名称': company, '股东信息': sh_str})
                    
                    df_sh_clean = pd.DataFrame(sh_grouped)
                    print(f"  股东信息整理完成: {len(df_sh_clean)} 家企业有股东数据")

                    # 合并到主数据
                    before_count = len(df_expanded)
                    df_expanded = df_expanded.merge(
                        df_sh_clean, on='企业名称', how='left', suffixes=('', '_sh')
                    )

                    # 处理股东列合并
                    target_col = '股东信息'
                    if target_col not in available_cols and target_col in df_expanded.columns:
                        available_cols.append(target_col)
                    # 处理带后缀的列
                    sh_col = f'{target_col}_sh'
                    if sh_col in df_expanded.columns:
                        if target_col in df_expanded.columns:
                            df_expanded[target_col] = df_expanded[target_col].fillna('').astype(str)
                            sh_vals = df_expanded[sh_col].fillna('').astype(str)
                            df_expanded[target_col] = df_expanded[target_col].where(
                                df_expanded[target_col] != '', sh_vals
                            )
                        else:
                            df_expanded = df_expanded.rename(columns={sh_col: target_col})
                            if target_col not in available_cols:
                                available_cols.append(target_col)
                        df_expanded = df_expanded.drop(columns=[sh_col])

                    matched = df_expanded[
                        df_expanded.get(target_col, pd.Series(dtype=str)).astype(str).str.strip() != ''
                    ].shape[0]
                    print(f"  股东信息合并完成: {matched}/{before_count} 条匹配到股东信息")
        except Exception as e:
            print(f"  股东信息文件处理失败: {e}，跳过合并")

    df_expanded[available_cols].to_csv(map_file, index=False, encoding='utf-8-sig')
    print(f"  映射文件已保存: {os.path.basename(map_file)}")

    # 生成支付宝核验文件（向量化，比iterrows快）
    print(f"\n正在生成支付宝核验文件...")
    random.seed(random_seed)
    
    # 向量化构造收款人姓名列
    if '法定代表人' in df_expanded.columns:
        names = df_expanded['法定代表人'].fillna('').astype(str).str.strip()
        names = names.mask((names == '') | (names == 'nan'), df_expanded['企业名称'].fillna('').astype(str).str.strip())
    else:
        names = df_expanded['企业名称'].fillna('').astype(str).str.strip()
    
    # 生成随机金额（向量化）
    import numpy as np
    np.random.seed(random_seed)
    amounts = np.round(np.random.uniform(1.0, 10.0, size=len(df_expanded)), 2)
    
    # 构建数据列表
    phones = df_expanded['手机号'].astype(str)
    alipay_rows = []
    for i, (phone, name, amount) in enumerate(zip(phones, names, amounts), 1):
        alipay_rows.append([i, phone, name, f"{amount:.2f}", ''])

    print(f"  支付宝核验记录: {len(alipay_rows)} 条")

    # 拆分文件
    parts = [alipay_rows[i:i + max_rows_per_file]
             for i in range(0, len(alipay_rows), max_rows_per_file)]
    print(f"  拆分为 {len(parts)} 份")

    output_files = []
    headers = ['序号（必填）', '收款方支付宝账号（必填）',
               '收款方姓名（必填）', '金额（必填，单位：元）', '备注（选填）']
    widths = [10, 22, 18, 18, 15]
    
    for idx, part in enumerate(parts):
        wb = Workbook()
        ws = wb.active
        ws.title = '工作表1'

        # 标题行
        ws.cell(1, 1, '【批量付钱】文件上传校验结果').font = Font(name='微软雅黑', size=11, bold=True)

        # 表头
        for c, h in enumerate(headers, 1):
            cell = ws.cell(2, c, h)
            cell.font = Font(name='微软雅黑', size=10, bold=True)
            cell.alignment = Alignment(horizontal='center')

        # 数据（使用append批量写入，比逐cell快）
        for row_data in part:
            ws.append(row_data)

        # 列宽
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(1, c).column_letter].width = w

        if len(parts) > 1:
            out_path = f'{out_base}_{idx + 1}.xlsx'
        else:
            out_path = f'{out_base}.xlsx'

        wb.save(out_path)
        output_files.append(out_path)
        print(f"    第{idx + 1}份: {len(part)} 行 -> {os.path.basename(out_path)}")

    return {
        'total_companies': len(df),
        'total_phones': len(alipay_rows),
        'num_files': len(parts),
        'output_files': output_files,
        'map_file': map_file,
        'skipped_verified': skipped_count,
        'industrial_filtered': industrial_filtered,
        'company_filtered': company_filtered,
    }


if __name__ == '__main__':
    # 测试用
    import sys
    if len(sys.argv) >= 3:
        result = generate_alipay_files(sys.argv[1], sys.argv[2])
        print(f"\n完成! 共 {result['total_phones']} 条记录, {result['num_files']} 份文件")

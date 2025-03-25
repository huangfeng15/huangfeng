import pdfplumber
import re
from typing import List, Dict

class PDFProcessor:
    def __init__(self, read_order: str, allow_empty: bool = False, custom_keys: List[str] = None):
        self.read_order = read_order
        self.allow_empty = allow_empty
        # 预处理键名：移除空白字符并标准化
        self.custom_keys = []
        self.original_keys = []
        if custom_keys:
            for key in custom_keys:
                if key.strip():
                    self.original_keys.append(key.strip())
                    normalized_key = self._normalize_text(key)
                    base_key = normalized_key.replace('(元)', '').replace('（元）', '')
                    self.custom_keys.append(base_key)
    
    def _normalize_text(self, text: str) -> str:
        """标准化文本，但保留更多原始格式"""
        if not text or not isinstance(text, str):
            return ""
        # 保留括号内的内容，只处理空白字符
        text = re.sub(r'\s+', '', text)
        # 移除中英文冒号
        text = text.rstrip('：:')
        # 替换全角字符为半角字符
        text = text.replace('（', '(').replace('）', ')')
        return text.lower()

    def _normalize_value(self, value: str, key: str = "") -> str:
        """标准化值，处理数字格式，根据键名进行特殊处理"""
        if not value or not isinstance(value, str):
            return ""
            
        # 对价格类键名进行特殊处理
        if any(price_key in key for price_key in ['控制价', '预算', '金额', '报价', '上限价']):
            return self._extract_price(value)
            
        # 其他值的常规处理
        try:
            # 尝试解析为数字（适用于纯数字值）
            value = re.sub(r',', '', value)
            value = "{:.2f}".format(float(value))
        except ValueError:
            pass
            
        return value
        
    def _extract_price(self, value: str) -> str:
        """从文本中提取价格，保留数字、逗号和小数点，并处理带单位和标题的情况"""
        if not value or not isinstance(value, str):
            return ""
        
        # 首先尝试从文本中分离出纯数字部分
        # 处理形如"采购上限价 533 333.33"的情况
        number_pattern = r'[\d\s,.]+$'  # 匹配尾部的数字、空格、逗号和小数点
        number_matches = re.search(number_pattern, value)
        
        if number_matches:
            # 提取匹配到的数字部分
            number_str = number_matches.group().strip()
            
            # 移除所有空格
            number_str = re.sub(r'\s+', '', number_str)
            
            # 处理可能的多个小数点
            parts = number_str.split('.')
            if len(parts) > 1:
                # 保留第一个小数点前的部分和第一个小数点后的部分
                number_str = parts[0] + '.' + parts[1]
                
            return number_str
        else:
            # 如果上面的模式没有匹配到，尝试提取任何数字序列
            price_pattern = r'[\d,.]+' 
            matches = re.findall(price_pattern, value)
            
            if matches:
                # 合并所有匹配结果
                extracted = ''.join(matches)
                
                # 处理可能的多个小数点
                parts = extracted.split('.')
                if len(parts) > 1:
                    # 保留第一个小数点前的部分和第一个小数点后的部分
                    extracted = parts[0] + '.' + parts[1]
                    
                return extracted
                
        return ""

    def _process_table(self, table: List[List]) -> List[Dict]:
        """处理表格数据，确保完整提取单元格内所有信息"""
        results = []
        if not table or not self.custom_keys:
            return results
        
        # 清理表格数据，但保留单元格内的空格和格式
        cleaned_table = []
        for row in table:
            if row and any(cell is not None and str(cell).strip() for cell in row):
                # 处理每个单元格，保留内部格式和间隔
                cleaned_row = []
                for cell in row:
                    if cell is None:
                        cleaned_row.append("")
                    else:
                        # 将单元格内容转换为字符串，保留间隔
                        cell_str = str(cell)
                        # 只清理前后空格，保留内部空格和格式
                        cell_str = cell_str.strip()
                        cleaned_row.append(cell_str)
                cleaned_table.append(cleaned_row)
        
        if not cleaned_table:
            return results
            
        # 根据阅读顺序处理
        if self.read_order == "left_to_right":
            results.extend(self._process_horizontal(cleaned_table))
        else:
            results.extend(self._process_vertical(cleaned_table))
            
        return results

    def _process_horizontal(self, table: List[List]) -> List[Dict]:
        """从左到右处理表格"""
        results = []
        found_keys = set()  # 跟踪已找到的键
        
        # 严格匹配自定义键
        for row in table:
            for i in range(len(row)):
                current_cell = row[i]
                if not current_cell:
                    continue
                
                current_normalized = self._normalize_text(current_cell)
                
                # 严格匹配自定义键
                matched_key = None
                for idx, key_template in enumerate(self.custom_keys):
                    # 完全相等才算匹配成功，避免"时间"匹配到"公示开始时间"等情况
                    if current_normalized == key_template:
                        matched_key = self.original_keys[idx]
                        found_keys.add(key_template)
                        break
                        
                if matched_key:
                    # 提取值：扫描右侧所有单元格，完整保留所有内容
                    for j in range(i + 1, len(row)):
                        next_cell = row[j]
                        if next_cell:  # 找到非空单元格
                            if next_cell or self.allow_empty:
                                results.append({
                                    'key': matched_key,
                                    'value': next_cell  # 直接使用完整的单元格内容
                                })
                            break
        
        # 第二步：宽松匹配未找到的键
        if len(found_keys) < len(self.custom_keys):
            for row in table:
                for i in range(len(row)):
                    current_cell = row[i]
                    if not current_cell:
                        continue
                    
                    current_normalized = self._normalize_text(current_cell)
                    
                    # 对于时间相关的键名，需要更严格的匹配规则
                    for idx, key_template in enumerate(self.custom_keys):
                        if key_template in found_keys:
                            continue  # 跳过已找到的键
                        
                        # 时间相关键名需要特殊处理
                        if '时间' in key_template:
                            if current_normalized == '时间' and key_template != '时间':
                                continue
                            
                            if not (key_template == current_normalized or 
                                   current_normalized.startswith(key_template)):
                                continue
                        else:
                            # 非时间类键名可以使用宽松匹配
                            if not (key_template in current_normalized or 
                                   current_normalized in key_template):
                                continue
                        
                        # 提取值：完整保留单元格内容
                        for j in range(i + 1, len(row)):
                            next_cell = row[j]
                            if next_cell:  # 找到非空单元格
                                if next_cell or self.allow_empty:
                                    results.append({
                                        'key': self.original_keys[idx],
                                        'value': next_cell  # 直接使用完整的单元格内容
                                    })
                                found_keys.add(key_template)
                                break
                        break

        return results

    def _process_vertical(self, table: List[List]) -> List[Dict]:
        """从上到下处理表格"""
        results = []
        found_keys = set()  # 跟踪已找到的键
        
        cols = len(table[0])
        rows = len(table)
        
        # 第一步：严格匹配自定义键
        for col in range(cols):
            for row in range(rows - 1):  # -1 确保有下一行
                current_cell = table[row][col]
                if not current_cell:
                    continue
                
                current_normalized = self._normalize_text(current_cell)
                
                # 严格匹配自定义键
                matched_key = None
                for idx, key_template in enumerate(self.custom_keys):
                    # 完全相等才算匹配成功
                    if current_normalized == key_template:
                        matched_key = self.original_keys[idx]
                        found_keys.add(key_template)
                        break
                        
                if matched_key:
                    # 寻找值：扫描下方所有单元格直到找到非空值
                    for next_row in range(row + 1, rows):
                        next_cell = table[next_row][col]
                        if next_cell:  # 找到非空单元格
                            if next_cell or self.allow_empty:
                                results.append({
                                    'key': matched_key,
                                    'value': next_cell  # 直接使用完整的单元格内容
                                })
                            break
        
        # 第二步：宽松匹配未找到的键
        if len(found_keys) < len(self.custom_keys):
            for col in range(cols):
                for row in range(rows - 1):
                    current_cell = table[row][col]
                    if not current_cell:
                        continue
                    
                    current_normalized = self._normalize_text(current_cell)
                    
                    for idx, key_template in enumerate(self.custom_keys):
                        if key_template in found_keys:
                            continue  # 跳过已找到的键
                        
                        # 时间相关键名需要特殊处理
                        if '时间' in key_template:
                            if current_normalized == '时间' and key_template != '时间':
                                continue
                            
                            if not (key_template == current_normalized or 
                                   current_normalized.startswith(key_template)):
                                continue
                        else:
                            # 非时间键名允许更宽松的匹配
                            if not (key_template in current_normalized or 
                                   current_normalized in key_template):
                                continue
                        
                        # 寻找值：完整保留单元格内容
                        for next_row in range(row + 1, rows):
                            next_cell = table[next_row][col]
                            if next_cell:  # 找到非空单元格
                                if next_cell or self.allow_empty:
                                    results.append({
                                        'key': self.original_keys[idx],
                                        'value': next_cell  # 直接使用完整的单元格内容
                                    })
                                found_keys.add(key_template)
                                break
                        break

        return results

    def _extract_date(self, value: str) -> str:
        """从时间字符串中提取日期部分"""
        if not value or not isinstance(value, str):
            return value
            
        # 常见日期格式的正则表达式模式
        date_patterns = [
            (r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日]?(?=\s|$)', lambda x: x),  # 2023-01-01 或 2023年01月01日
            (r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日]?\s+\d{1,2}[:：时]\d{1,2}', lambda x: x.split()[0]),  # 2023-01-01 10:30
            (r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日]?.*\d{1,2}[:：时]\d{1,2}[:：分]\d{1,2}', lambda x: x.split()[0])  # 2023-01-01 10:30:00
        ]
        
        for pattern, extractor in date_patterns:
            match = re.search(pattern, value)
            if match:
                return extractor(match.group())
        
        return value

    def _deduplicate_results(self, results: List[Dict]) -> List[Dict]:
        """优化的去重逻辑，合并相同值，避免错误匹配"""
        final_results = []
        seen_keys = {}
        
        # 额外处理时间相关键名的映射
        time_related_keys = {}
        for idx, key in enumerate(self.original_keys):
            if '时间' in self._normalize_text(key):
                key_base = self._normalize_text(key)
                time_related_keys[key_base] = idx
        
        # 按照预定义键名的顺序处理
        for idx, template_key in enumerate(self.custom_keys):
            original_key = self.original_keys[idx]
            template_base = template_key.split('(')[0].split('（')[0]
            is_time_key = '时间' in template_key
            is_price_key = any(x in original_key for x in ['控制价', '预算', '金额', '报价'])
            
            best_match = None
            best_value = ""
            
            # 查找最佳匹配
            for result in results:
                key_normalized = self._normalize_text(result['key'])
                
                # 处理时间键名
                if is_time_key:
                    # 如果当前结果的键是"时间"而当前模板不是纯"时间"，则跳过
                    if key_normalized == '时间' and template_key != '时间':
                        continue
                    
                    # 对于时间键名，只有精确匹配或以模板开头才接受
                    if not (template_key == key_normalized or
                           key_normalized.startswith(template_key)):
                        continue
                else:
                    # 非时间键名允许更宽松的匹配
                    if not (template_base in key_normalized or key_normalized in template_base):
                        continue
                
                # 优先选择有值的结果
                current_value = result['value'].strip()
                if not best_match or (current_value and not best_value):
                    best_match = result
                    best_value = current_value
            
            # 如果找到匹配且还未添加过
            if best_match:
                # 更新键名为预定义的键名，确保输出一致性
                best_match['key'] = original_key
                
                # 处理时间值：如果是时间相关的键名，只保留日期部分
                if is_time_key:
                    best_match['value'] = self._extract_date(best_match['value'])
                # 如果是价格类键名，进行特殊处理
                elif is_price_key:
                    best_match['value'] = self._extract_price(best_match['value'])
                
                # 检查是否已有相同键名
                key_base = self._normalize_text(original_key)
                if key_base not in seen_keys:
                    seen_keys[key_base] = True
                    final_results.append(best_match)
        
        return final_results

    def process_pdf(self, file_path: str) -> List[Dict]:
        """处理PDF文件"""
        all_results = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    # 处理表格
                    tables = page.extract_tables()
                    for table in tables:
                        results = self._process_table(table)
                        if results:
                            all_results.extend(results)
                            
                    # 启用文本块处理，补充表格提取无法识别的部分
                    text_blocks = self._extract_text_blocks(page)
                    if text_blocks:
                        all_results.extend(self._process_text_blocks(text_blocks))
                            
        except Exception as e:
            raise Exception(f"PDF处理错误: {str(e)}")
            
        return self._deduplicate_results(all_results)

    def _extract_text_blocks(self, page) -> List[Dict]:
        try:
            # 大幅增大x容差值，以便能够正确处理单元格内的大空格分隔
            words = page.extract_words(keep_blank_chars=True, x_tolerance=15, y_tolerance=5)
            
            if not words:
                return []
                
            # 确保所有必需的键都存在
            words = [w for w in words if all(k in w for k in ['x0', 'top', 'text'])]
            
            if self.read_order == "top_to_bottom":
                words.sort(key=lambda x: (x['top'], x['x0']))
            else:
                words.sort(key=lambda x: (x['x0'], x['top']))
                
            blocks = []
            current_block = ""
            current_position = None
            current_top = None
            last_x1 = None  # 跟踪上一个单词的结束位置
            
            for word in words:
                if not current_position:
                    current_position = (word['x0'], word['top'])
                    current_top = word['top']
                    current_block = word['text']
                    last_x1 = word.get('x1', word['x0'] + len(word['text']) * 5)  # 估算结束位置
                elif abs(word['top'] - current_top) < 8:  # 同一行
                    space_needed = True
                    
                    # 判断是否需要添加空格 - 基于实际间距
                    if last_x1 and 'x0' in word:
                        # 如果间距大于阈值，添加空格，保留原始格式
                        x_gap = word['x0'] - last_x1
                        space_needed = x_gap > 10  # 根据实际PDF格式调整
                        
                    # 添加适当的连接符
                    if space_needed and current_block and not current_block.endswith(' ') and not word['text'].startswith(' '):
                        current_block += " " + word['text']
                    else:
                        current_block += word['text']
                        
                    last_x1 = word.get('x1', word['x0'] + len(word['text']) * 5)
                else:
                    if current_block.strip():
                        blocks.append({
                            'text': current_block.strip(),
                            'position': current_position
                        })
                    current_block = word['text']
                    current_position = (word['x0'], word['top'])
                    current_top = word['top']
                    last_x1 = word.get('x1', word['x0'] + len(word['text']) * 5)
            
            if current_block and current_block.strip():
                blocks.append({
                    'text': current_block.strip(),
                    'position': current_position
                })
                
            return blocks
            
        except Exception as e:
            raise Exception(f"文本块提取错误: {str(e)}")

    def _is_key(self, text: str) -> bool:
        """改进的键名匹配逻辑"""
        if not text or not self.custom_keys:
            return False
            
        # 标准化文本，但保留括号等特殊字符
        test_text = self._normalize_text(text)
        
        for key in self.custom_keys:
            # 处理可能的变体
            variants = [
                key,  # 原始键名
                key + ":",  # 英文冒号
                key + "：",  # 中文冒号
                key + "(元)",  # 带单位
                key + "（元）",  # 带单位（中文括号）
                key.replace("（元）", "").replace("(元)", "")  # 去掉单位的版本
            ]
            
            # 任何一个变体匹配就返回True
            if any(variant in test_text for variant in variants):
                return True
                
        return False

    def _process_text_blocks(self, blocks: List[Dict]) -> List[Dict]:
        results = []
        
        # 首先检查文本块中是否包含冒号分隔的键值对
        for i, block in enumerate(blocks):
            current_text = block['text']
            
            # 检查是否包含中文或英文冒号
            if '：' in current_text or ':' in current_text:
                # 分割文本获取键值对
                separator = '：' if '：' in current_text else ':'
                parts = current_text.split(separator, 1)
                
                if len(parts) == 2:
                    key_part = parts[0].strip()
                    value_part = parts[1].strip()
                    
                    # 检查键名是否在自定义键列表中
                    if self._is_key(key_part):
                        results.append({
                            'key': key_part,
                            'value': value_part if value_part or self.allow_empty else ""
                        })
        
        # 然后处理相邻的文本块作为可能的键值对
        for i in range(len(blocks) - 1):
            current_block = blocks[i]['text']
            next_block = blocks[i+1]['text']
            
            # 已经在上面处理过的带冒号的块就跳过
            if ':' in current_block or '：' in current_block:
                continue
                
            # 检查当前块是否为键名
            if self._is_key(current_block):
                # 如果下一个块也是键名，则跳过
                if self._is_key(next_block):
                    continue
                    
                results.append({
                    'key': current_block.rstrip('：:'),
                    'value': next_block if next_block or self.allow_empty else ""
                })
        
        return results

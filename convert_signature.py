# -*- coding: utf-8 -*-
"""
公共危險物品場所自主檢查表 - 簽章欄自動轉換程式 (v2.2)
將舊版簽章表格（管理權人、檢查及會同人員、審核人員）
自動轉換為新版格式（檢 查 日 期、保 安 檢 查 員 簽 名、保 安 監 督 人 複 核 簽 名）
樣式規範：
- 完整 3 欄邊框（頂、底、左、右與中隔線）
- 文字靠左對齊，預留充足手寫日期、簽名與蓋章空間
- 充足行高 (trHeight 700 dxa)
- 僅更動簽章列，其餘內容、字型、排版百分之百保留。
"""

import os
import sys
import glob
import shutil
import zipfile
import xml.etree.ElementTree as ET

# 註冊所有常見 Word XML 命名空間
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

W_NS = f"{{{NAMESPACES['w']}}}"
NS = {'w': NAMESPACES['w']}


def create_new_signature_row(labels_with_specs):
    """
    建立新版簽章列 (tr)
    labels_with_specs: list of (text, width_dxa, grid_span, is_first, is_last)
    """
    tr = ET.Element(f"{W_NS}tr")
    
    # 列屬性：充足行高以供手寫簽名蓋章
    trPr = ET.SubElement(tr, f"{W_NS}trPr")
    ET.SubElement(trPr, f"{W_NS}cantSplit")
    ET.SubElement(trPr, f"{W_NS}trHeight", {f"{W_NS}val": "700"})
    ET.SubElement(trPr, f"{W_NS}jc", {f"{W_NS}val": "center"})
    
    total_cols = len(labels_with_specs)
    
    for idx, (label, w, gs) in enumerate(labels_with_specs):
        is_first = (idx == 0)
        is_last = (idx == total_cols - 1)
        
        tc = ET.SubElement(tr, f"{W_NS}tc")
        tcPr = ET.SubElement(tc, f"{W_NS}tcPr")
        ET.SubElement(tcPr, f"{W_NS}tcW", {f"{W_NS}w": str(w), f"{W_NS}type": "dxa"})
        if gs > 1:
            ET.SubElement(tcPr, f"{W_NS}gridSpan", {f"{W_NS}val": str(gs)})
            
        # 邊界設定：完整四邊框線，外邊框粗線(12)，內部分隔線細線(4/8)
        tcBorders = ET.SubElement(tcPr, f"{W_NS}tcBorders")
        ET.SubElement(tcBorders, f"{W_NS}top", {
            f"{W_NS}val": "single",
            f"{W_NS}sz": "12",
            f"{W_NS}space": "0",
            f"{W_NS}color": "auto"
        })
        ET.SubElement(tcBorders, f"{W_NS}bottom", {
            f"{W_NS}val": "single",
            f"{W_NS}sz": "12",
            f"{W_NS}space": "0",
            f"{W_NS}color": "auto"
        })
        ET.SubElement(tcBorders, f"{W_NS}left", {
            f"{W_NS}val": "single",
            f"{W_NS}sz": "12" if is_first else "4",
            f"{W_NS}space": "0",
            f"{W_NS}color": "auto"
        })
        ET.SubElement(tcBorders, f"{W_NS}right", {
            f"{W_NS}val": "single",
            f"{W_NS}sz": "12" if is_last else "4",
            f"{W_NS}space": "0",
            f"{W_NS}color": "auto"
        })
        
        # 垂直置中
        ET.SubElement(tcPr, f"{W_NS}vAlign", {f"{W_NS}val": "center"})
        
        # 段落與文字（靠左對齊，留給後方簽名蓋章空間）
        p = ET.SubElement(tc, f"{W_NS}p")
        pPr = ET.SubElement(p, f"{W_NS}pPr")
        ET.SubElement(pPr, f"{W_NS}spacing", {f"{W_NS}line": "280", f"{W_NS}lineRule": "exact"})
        ET.SubElement(pPr, f"{W_NS}jc", {f"{W_NS}val": "left"})
        ET.SubElement(pPr, f"{W_NS}ind", {f"{W_NS}left": "100"})  # 靠左預留適當邊距
        
        rPr_p = ET.SubElement(pPr, f"{W_NS}rPr")
        ET.SubElement(rPr_p, f"{W_NS}rFonts", {
            f"{W_NS}ascii": "標楷體",
            f"{W_NS}eastAsia": "標楷體",
            f"{W_NS}hAnsi": "標楷體"
        })
        ET.SubElement(rPr_p, f"{W_NS}sz", {f"{W_NS}val": "20"})  # 10pt (更小字體，留出最大簽章空間)
        
        r = ET.SubElement(p, f"{W_NS}r")
        rPr = ET.SubElement(r, f"{W_NS}rPr")
        ET.SubElement(rPr, f"{W_NS}rFonts", {
            f"{W_NS}ascii": "標楷體",
            f"{W_NS}eastAsia": "標楷體",
            f"{W_NS}hAnsi": "標楷體",
            f"{W_NS}hint": "eastAsia"
        })
        ET.SubElement(rPr, f"{W_NS}sz", {f"{W_NS}val": "20"})  # 10pt
        
        t = ET.SubElement(r, f"{W_NS}t")
        t.text = label
        
    return tr


def convert_docx_file(src_path, dst_path):
    """
    轉換單一 docx 檔案：
    1. 若開頭有「高雄市公共危險物品場所聯合檢查表」，刪除該頁直到「附件4」開頭。
    2. 刪除表格中的「備註｜抽查 年「 」月自主檢查表」列。
    3. 僅修改舊版簽章列為新版三欄簽章列，保留其他所有內容。
    傳回 (成功與否, 統計資訊 dict)
    """
    if not os.path.exists(src_path):
        return False, {'signatures': 0, 'removed_joint': False, 'removed_choucha': 0}
        
    os.makedirs(os.path.dirname(os.path.abspath(dst_path)), exist_ok=True)
    
    with zipfile.ZipFile(src_path, 'r') as zin:
        xml_content = zin.read('word/document.xml')
        root = ET.fromstring(xml_content)
        
        removed_joint = False
        removed_choucha = 0
        modified_count = 0
        
        # 1. 偵測並移除第一頁「聯合檢查表」（從頭刪到「附件4」之前）
        body = root.find('./w:body', NS)
        if body is not None:
            children = list(body)
            joint_idx = -1
            att4_idx = -1
            for idx, child in enumerate(children[:15]):
                txt = ''.join(t.text for t in child.findall('.//w:t', NS) if t.text).strip()
                if '聯合檢查表' in txt and joint_idx == -1:
                    joint_idx = idx
                if ('附件4' in txt or '附件４' in txt or '消防安全檢查紀錄表' in txt) and att4_idx == -1:
                    att4_idx = idx
            
            if joint_idx != -1 and att4_idx > joint_idx:
                for el in children[:att4_idx]:
                    body.remove(el)
                # 若附件4段落本身有 pageBreakBefore，一併移除避免開頭空白頁
                att4_el = children[att4_idx]
                pPr = att4_el.find('./w:pPr', NS)
                if pPr is not None:
                    pbb = pPr.find('./w:pageBreakBefore', NS)
                    if pbb is not None:
                        pPr.remove(pbb)
                removed_joint = True
        
        tables = root.findall('.//w:tbl', NS)
        
        for tbl in tables:
            # 2. 刪除「備註｜抽查 年「 」月自主檢查表」列
            rows = tbl.findall('./w:tr', NS)
            for row in rows:
                text = ''.join([t.text for t in row.findall('.//w:t', NS) if t.text])
                if '抽查' in text and '自主檢查表' in text:
                    tbl.remove(row)
                    removed_choucha += 1

            # 3. 替換簽章列
            rows = tbl.findall('./w:tr', NS)
            rows_to_replace = []
            
            for r_idx, row in enumerate(rows):
                text = ''.join([t.text for t in row.findall('.//w:t', NS) if t.text])
                if '管理權人或現場有關人員' in text and ('簽章' in text or '檢查及會同人員' in text or '審核人員' in text):
                    rows_to_replace.append(r_idx)
            
            for r_idx in reversed(rows_to_replace):
                target_row = rows[r_idx]
                
                # 動態精準計算最佳比例：檢查日期約 20%~25%（窄欄），其餘約 75%~80% 均分給兩位簽署人
                grid = tbl.find('./w:tblGrid', NS)
                col_widths = []
                if grid is not None:
                    col_widths = [int(c.attrib.get(f"{W_NS}w", '0')) for c in grid.findall('./w:gridCol', NS)]
                    
                N = len(col_widths)
                total_w = sum(col_widths) if col_widths else 0
                
                if N >= 3 and total_w > 0:
                    target1 = 0.22 * total_w
                    target2 = 0.39 * total_w
                    target3 = 0.39 * total_w
                    
                    best_score = float('inf')
                    best_ij = (1, 2)
                    
                    for i in range(1, N - 1):
                        w1 = sum(col_widths[:i])
                        if w1 < 1200:
                            continue
                        for j in range(i + 1, N):
                            w2 = sum(col_widths[i:j])
                            w3 = sum(col_widths[j:])
                            if w2 < 1500 or w3 < 1500:
                                continue
                                
                            score = abs(w1 - target1) * 3.0 + abs(w2 - target2) + abs(w3 - target3) + abs(w2 - w3) * 0.5
                            if score < best_score:
                                best_score = score
                                best_ij = (i, j)
                                
                    i, j = best_ij
                    w1 = sum(col_widths[:i])
                    w2 = sum(col_widths[i:j])
                    w3 = sum(col_widths[j:])
                    gs1, gs2, gs3 = i, j - i, N - j
                else:
                    cells = target_row.findall('./w:tc', NS)
                    def get_w(c):
                        w_el = c.find('.//w:tcW', NS)
                        return int(w_el.attrib.get(f"{W_NS}w", '0')) if w_el is not None else 0
                    def get_gs(c):
                        gs_el = c.find('.//w:gridSpan', NS)
                        return int(gs_el.attrib.get(f"{W_NS}val", '1')) if gs_el is not None else 1
                        
                    tot = sum(get_w(c) for c in cells) or 10340
                    tot_gs = sum(get_gs(c) for c in cells) or 6
                    w1 = int(tot * 0.22)
                    w2 = int((tot - w1) / 2)
                    w3 = tot - w1 - w2
                    gs1 = max(1, round(tot_gs * 0.22))
                    gs2 = max(1, round((tot_gs - gs1) / 2))
                    gs3 = max(1, tot_gs - gs1 - gs2)
                
                specs = [
                    ('檢查日期：', w1, gs1),
                    ('保安檢查員簽名：', w2, gs2),
                    ('保安監督人複核簽名：', w3, gs3)
                ]
                
                new_tr = create_new_signature_row(specs)
                
                row_elements = list(tbl)
                target_pos = row_elements.index(target_row)
                
                # 移除原空白簽名手寫列
                if target_pos + 1 < len(row_elements) and row_elements[target_pos + 1].tag.endswith('tr'):
                    tbl.remove(row_elements[target_pos + 1])
                    
                tbl.remove(target_row)
                tbl.insert(target_pos, new_tr)
                modified_count += 1
                
        new_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        
    temp_dst = dst_path + '.tmp'
    with zipfile.ZipFile(src_path, 'r') as zin, zipfile.ZipFile(temp_dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))
                
    if os.path.exists(dst_path):
        os.remove(dst_path)
    os.rename(temp_dst, dst_path)
    
    return True, {
        'signatures': modified_count,
        'removed_joint': removed_joint,
        'removed_choucha': removed_choucha
    }


def process_targets(targets, output_folder=None):
    files_to_process = []
    for t in targets:
        t = t.strip('"').strip("'")
        if not os.path.exists(t):
            continue
        if os.path.isfile(t):
            if t.lower().endswith('.docx') and not os.path.basename(t).startswith('~$'):
                files_to_process.append(t)
        elif os.path.isdir(t):
            for root_dir, _, files in os.walk(t):
                for f in files:
                    if f.lower().endswith('.docx') and not f.startswith('~$'):
                        files_to_process.append(os.path.join(root_dir, f))
                        
    if not files_to_process:
        print("[提示] 沒有找到需要轉換的 .docx 檔案。")
        return
        
    print("=" * 65)
    print(f"  公共危險物品場所自主檢查表 - 自動轉換程式 (v2.3)")
    print(f"  待處理檔案數量：{len(files_to_process)} 個")
    print("=" * 65)
    
    success_count = 0
    total_tables = 0
    total_joint = 0
    total_choucha = 0
    
    for i, src_file in enumerate(files_to_process, 1):
        file_dir = os.path.dirname(os.path.abspath(src_file))
        file_name = os.path.basename(src_file)
        
        if output_folder:
            dest_dir = os.path.abspath(output_folder)
        else:
            dest_dir = os.path.join(file_dir, "轉換完成")
            
        dst_file = os.path.join(dest_dir, file_name)
        
        try:
            ok, stats = convert_docx_file(src_file, dst_file)
            if ok:
                success_count += 1
                sig_cnt = stats['signatures']
                total_tables += sig_cnt
                if stats['removed_joint']:
                    total_joint += 1
                total_choucha += stats['removed_choucha']
                tags = [f"簽章{sig_cnt:02d}處"]
                if stats['removed_joint']:
                    tags.append("刪聯合表")
                if stats['removed_choucha'] > 0:
                    tags.append(f"刪抽查列x{stats['removed_choucha']}")
                print(f"[{i:02d}/{len(files_to_process):02d}] 成功 ({' / '.join(tags)}): {file_name}")
            else:
                print(f"[{i:02d}/{len(files_to_process):02d}] 失敗: {file_name}")
        except Exception as e:
            print(f"[{i:02d}/{len(files_to_process):02d}] 發生錯誤: {file_name} -> {e}")
            
    print("=" * 65)
    print(f"  處理完畢！成功處理 {success_count} 個檔案：")
    print(f"  - 替換簽章表格：{total_tables} 處")
    print(f"  - 刪除聯合檢查表：{total_joint} 份")
    print(f"  - 刪除抽查自主檢查表備註列：{total_choucha} 列")
    print(f"  轉換後的檔案已儲存至：")
    if output_folder:
        print(f"  -> {os.path.abspath(output_folder)}")
    else:
        print(f"  -> 各來源檔案所在資料夾的「轉換完成」目錄下")
    print("=" * 65)


if __name__ == '__main__':
    args = sys.argv[1:]
    if not args:
        default_input = r"路竹30倍以上-16家ok"
        default_output = r"路竹30倍以上-轉換完成"
        
        if os.path.exists(default_input):
            print(f"[*] 偵測到預設資料夾「{default_input}」，即將開始批次轉換...")
            process_targets([default_input], output_folder=default_output)
        else:
            cur_docx = glob.glob("*.docx")
            if cur_docx:
                print(f"[*] 偵測到當前目錄下有 {len(cur_docx)} 個 .docx 檔案，即將開始轉換...")
                process_targets(cur_docx, output_folder="轉換完成")
            else:
                print("請將檔案或資料夾拖曳至本程式上，或傳入路徑參數。")
    else:
        out_folder = None
        filtered_args = []
        skip_next = False
        for i, a in enumerate(args):
            if skip_next:
                skip_next = False
                continue
            if a in ('-o', '--output', '--output-dir'):
                if i + 1 < len(args):
                    out_folder = args[i + 1]
                    skip_next = True
            else:
                filtered_args.append(a)
                
        process_targets(filtered_args, output_folder=out_folder)

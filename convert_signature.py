# -*- coding: utf-8 -*-
"""
公共危險物品場所自主檢查表 - 自動轉換程式 (v3.0 - 11308起適用範例版)
功能：
1. 自動偵測 30 倍以上 / 30 倍以下（依保安監督人是否免設或合計倍數，亦支援手動指定 --mode above|below）。
2. 自動從原檔「附件4 基本資料」擷取場所名稱、地址、電話、管理權人、保安監督人資訊。
3. 自動移除第一頁「聯合檢查表」與原「附件4（含抽查備註列）」，並於最前面取代插入對應版本之「附件10 公共危險物品保安監督自主檢查表」並自動回填廠商基本資料。
4. 刪除所有附表（附表三～附表十二）表格內之舊版 2 列簽章區塊，並於表格下方（備註行之後）改為範例樣式之黑色粗體 12pt 標楷體簽章行：「保安檢查員簽章：」（30倍以上）或「專人簽章：」（30倍以下）。
"""

import os
import re
import sys
import glob
import zipfile
from xml.sax.saxutils import escape as xml_escape
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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_ABOVE_PATH = os.path.join(SCRIPT_DIR, 'templates', 'att10_above.xml')
TEMPLATE_BELOW_PATH = os.path.join(SCRIPT_DIR, 'templates', 'att10_below.xml')

SUBTABLE_KEYWORDS = ('附表三', '附表四', '附表五', '附表六', '附表七', '附表八', '附表九', '附表十', '附表十一', '附表十二')


def get_text(elem):
    return ''.join(t.text for t in elem.findall('.//w:t', NS) if t.text)


def create_signature_paragraph(sig_text):
    """
    建立範例樣式的表格外簽章段落（黑色粗體 12pt 標楷體）
    """
    p = ET.Element(f"{W_NS}p")
    pPr = ET.SubElement(p, f"{W_NS}pPr")
    ET.SubElement(pPr, f"{W_NS}widowControl")
    ET.SubElement(pPr, f"{W_NS}spacing", {f"{W_NS}before": "60", f"{W_NS}after": "60"})
    rPr_p = ET.SubElement(pPr, f"{W_NS}rPr")
    ET.SubElement(rPr_p, f"{W_NS}rFonts", {
        f"{W_NS}ascii": "標楷體",
        f"{W_NS}eastAsia": "標楷體",
        f"{W_NS}hAnsi": "標楷體"
    })
    ET.SubElement(rPr_p, f"{W_NS}b")
    ET.SubElement(rPr_p, f"{W_NS}bCs")
    ET.SubElement(rPr_p, f"{W_NS}sz", {f"{W_NS}val": "24"})

    r = ET.SubElement(p, f"{W_NS}r")
    rPr = ET.SubElement(r, f"{W_NS}rPr")
    ET.SubElement(rPr, f"{W_NS}rFonts", {
        f"{W_NS}ascii": "標楷體",
        f"{W_NS}eastAsia": "標楷體",
        f"{W_NS}hAnsi": "標楷體",
        f"{W_NS}hint": "eastAsia"
    })
    ET.SubElement(rPr, f"{W_NS}b")
    ET.SubElement(rPr, f"{W_NS}bCs")
    ET.SubElement(rPr, f"{W_NS}color", {f"{W_NS}val": "000000"})
    ET.SubElement(rPr, f"{W_NS}sz", {f"{W_NS}val": "24"})  # 12pt
    t = ET.SubElement(r, f"{W_NS}t")
    t.text = sig_text
    return p


def extract_basic_info(body):
    """
    從「聯合檢查表」與「附件4 基本資料」擷取場所與人員資訊，並判斷是否為 30 倍以上。
    """
    info = {
        'SITE_NAME': '',
        'SITE_PHONE': '',
        'SITE_ADDR': '',
        'OWNER_NAME': '',
        'OWNER_PHONE': '',
        'OWNER_ID': '',
        'SUPERVISOR_NAME': '',
        'SUPERVISOR_PHONE': '',
        'SUPERVISOR_ID': '',
    }
    children = list(body)
    joint_name = ''
    joint_addr = ''
    att4_tbl_idx = -1

    for idx, ch in enumerate(children):
        if ch.tag != f"{W_NS}tbl":
            continue
        txt = get_text(ch)
        if '聯合檢查表' in txt and not joint_name:
            m_name = re.search(r'事業單位名稱[：:]\s*([^\s]+)', txt)
            if m_name:
                joint_name = m_name.group(1).strip()
            m_addr = re.search(r'事業單位地址[：:]\s*([^\s]+)', txt)
            if m_addr:
                joint_addr = m_addr.group(1).strip()
        if '一、基本資料' in txt and ('管理權人' in txt or '保安監督人' in txt) and att4_tbl_idx == -1:
            att4_tbl_idx = idx
            for row in ch.findall('./w:tr', NS):
                cells = [get_text(c).strip() for c in row.findall('./w:tc', NS)]
                if not cells:
                    continue
                key = cells[0].replace(' ', '').replace('\u3000', '')
                # 跳過「五、保安監督及安全管理」的勾選檢查列，只讀取「一、基本資料」區
                if len(cells) >= 2 and '符合規定' in cells[1]:
                    continue
                if key == '場所名稱' and len(cells) >= 2 and not info['SITE_NAME']:
                    info['SITE_NAME'] = cells[1]
                elif key == '場所地址' and len(cells) >= 2 and not info['SITE_ADDR']:
                    info['SITE_ADDR'] = cells[1]
                elif key == '管理權人' and len(cells) >= 2 and not info['OWNER_NAME']:
                    info['OWNER_NAME'] = cells[1]
                    if len(cells) > 3:
                        info['OWNER_ID'] = cells[3]
                    if len(cells) > 5:
                        info['OWNER_PHONE'] = cells[5]
                elif key == '保安監督人' and len(cells) >= 2 and not info['SUPERVISOR_NAME']:
                    info['SUPERVISOR_NAME'] = cells[1]
                    if len(cells) > 3:
                        info['SUPERVISOR_ID'] = cells[3]
                    if len(cells) > 5:
                        info['SUPERVISOR_PHONE'] = cells[5]

    if not info['SITE_NAME'] and joint_name:
        info['SITE_NAME'] = joint_name
    if not info['SITE_ADDR'] and joint_addr:
        info['SITE_ADDR'] = joint_addr
    info['SITE_PHONE'] = info['OWNER_PHONE'] or info['SUPERVISOR_PHONE']

    # 自動判定 30倍以上 vs 30倍以下
    is_above = None
    sup = info['SUPERVISOR_NAME'].strip()
    if sup:
        if '免設' in sup:
            is_above = False
        else:
            is_above = True

    # 若無附件4（例如獨立數量表），嘗試從附表十二合計倍數判定
    if is_above is None:
        for ch in children:
            if ch.tag == f"{W_NS}tbl":
                for row in ch.findall('./w:tr', NS):
                    cells = [get_text(c).strip() for c in row.findall('./w:tc', NS)]
                    for ci, val in enumerate(cells):
                        if '合計倍數' in val and ci + 1 < len(cells):
                            m_num = re.search(r'[\d.]+', cells[ci + 1])
                            if m_num:
                                try:
                                    is_above = float(m_num.group(0)) >= 30.0
                                except ValueError:
                                    pass
    if is_above is None:
        is_above = True

    return info, att4_tbl_idx, is_above


def load_att10_elements(is_above, info):
    tpl_path = TEMPLATE_ABOVE_PATH if is_above else TEMPLATE_BELOW_PATH
    with open(tpl_path, 'r', encoding='utf-8') as f:
        xml_str = f.read()
    for k, v in info.items():
        xml_str = xml_str.replace(f'{{{{{k}}}}}', xml_escape(v or ''))
    wrapper = ET.fromstring(xml_str)
    return list(wrapper)


def convert_docx_file(src_path, dst_path, mode='auto'):
    """
    轉換單一 docx 檔案 (v3.0)
    mode: 'auto' | 'above' | 'below'
    """
    if not os.path.exists(src_path):
        return False, {}

    os.makedirs(os.path.dirname(os.path.abspath(dst_path)), exist_ok=True)

    with zipfile.ZipFile(src_path, 'r') as zin:
        xml_content = zin.read('word/document.xml')
        root = ET.fromstring(xml_content)

        body = root.find('./w:body', NS)
        if body is None:
            return False, {}

        info, att4_tbl_idx, detected_above = extract_basic_info(body)
        if mode == 'above':
            is_above = True
        elif mode == 'below':
            is_above = False
        else:
            is_above = detected_above

        sig_label = '保安檢查員簽章：' if is_above else '專人簽章：'
        inserted_att10 = False
        removed_choucha = 0
        modified_sigs = 0

        # 1. 若有附件4主表，將「聯合檢查表 + 附件4 + 附件4備註」全部取代為「附件10」
        att10_tbl_ref = None
        if att4_tbl_idx != -1:
            children = list(body)
            first_sub_idx = -1
            for idx in range(att4_tbl_idx + 1, len(children)):
                ch = children[idx]
                if ch.tag == f"{W_NS}sectPr":
                    first_sub_idx = idx
                    break
                if ch.tag == f"{W_NS}tbl":
                    first_sub_idx = idx
                    break
                if ch.tag == f"{W_NS}p":
                    txt = get_text(ch).strip()
                    if any(kw in txt for kw in SUBTABLE_KEYWORDS):
                        if not txt.startswith(('一、', '二、', '備註')) and '附表三至附表十一' not in txt and '1場所1表' not in txt:
                            first_sub_idx = idx
                            break

            if first_sub_idx == -1:
                first_sub_idx = len(children)

            for el in children[:first_sub_idx]:
                body.remove(el)

            # 若第一個附表標題本身帶有 pageBreakBefore，移除避免多出空白頁
            if first_sub_idx < len(children) and children[first_sub_idx].tag == f"{W_NS}p":
                pPr = children[first_sub_idx].find('./w:pPr', NS)
                if pPr is not None:
                    pbb = pPr.find('./w:pageBreakBefore', NS)
                    if pbb is not None:
                        pPr.remove(pbb)

            att10_elems = load_att10_elements(is_above, info)
            for insert_pos, el in enumerate(att10_elems):
                body.insert(insert_pos, el)
                if el.tag == f"{W_NS}tbl":
                    att10_tbl_ref = el
            inserted_att10 = True

        # 2. 處理所有附表（附表三～附表十二）：刪除抽查列、刪除舊簽章 2 列、表格下方插入範例簽章段落
        body_children = list(body)
        # 由後往前處理，避免 insert 影響索引
        for idx in range(len(body_children) - 1, -1, -1):
            el = body_children[idx]
            if el.tag != f"{W_NS}tbl":
                continue
            if el is att10_tbl_ref:
                continue

            # 刪除抽查備註列（防呆）
            for row in list(el.findall('./w:tr', NS)):
                rtxt = get_text(row)
                if '抽查' in rtxt and '自主檢查表' in rtxt:
                    el.remove(row)
                    removed_choucha += 1

            # 刪除舊簽章列（管理權人或現場有關人員 + 下一列空白簽名列）
            rows = el.findall('./w:tr', NS)
            rows_to_remove = []
            for r_idx, row in enumerate(rows):
                rtxt = get_text(row)
                if '管理權人或現場有關人員' in rtxt and ('簽章' in rtxt or '檢查及會同人員' in rtxt or '審核人員' in rtxt):
                    rows_to_remove.append(r_idx)

            removed_old_sig = False
            for r_idx in reversed(rows_to_remove):
                target_row = rows[r_idx]
                row_elements = list(el)
                pos = row_elements.index(target_row)
                if pos + 1 < len(row_elements) and row_elements[pos + 1].tag == f"{W_NS}tr":
                    el.remove(row_elements[pos + 1])
                el.remove(target_row)
                removed_old_sig = True

            if removed_old_sig:
                # 尋找插入點：若表格後面有「※本表欄位所稱辦法...」段落，插在其後；否則緊接在表格之後
                cur_children = list(body)
                tbl_pos = cur_children.index(el)
                insert_after = tbl_pos

                # 往後看最多 2 個段落尋找 ※備註行
                for look in range(tbl_pos + 1, min(tbl_pos + 3, len(cur_children))):
                    nxt = cur_children[look]
                    if nxt.tag != f"{W_NS}p":
                        break
                    ntxt = get_text(nxt).strip()
                    if '本表欄位所稱辦法' in ntxt or ntxt.startswith('※'):
                        insert_after = look
                        # 移除 ※ 段落內原本藏著的手動分頁符號 (<w:br w:type="page"/>)，避免把簽章推到下一頁
                        for r_el in list(nxt.findall('./w:r', NS)):
                            for br_el in list(r_el.findall('./w:br', NS)):
                                if br_el.attrib.get(f"{W_NS}type") == 'page':
                                    r_el.remove(br_el)
                        # 設定 keepNext 確保備註與簽章留在同一頁
                        nxt_pPr = nxt.find('./w:pPr', NS)
                        if nxt_pPr is None:
                            nxt_pPr = ET.Element(f"{W_NS}pPr")
                            nxt.insert(0, nxt_pPr)
                        if nxt_pPr.find('./w:keepNext', NS) is None:
                            ET.SubElement(nxt_pPr, f"{W_NS}keepNext")
                        break
                    elif ntxt != '':
                        break

                # 檢查是否已經有簽章段落避免重複
                already_has = False
                for look in range(insert_after + 1, min(insert_after + 3, len(cur_children))):
                    nxt = cur_children[look]
                    if nxt.tag == f"{W_NS}p":
                        ntxt = get_text(nxt).strip()
                        if '保安檢查員簽章' in ntxt or '專人簽章' in ntxt:
                            already_has = True
                            break
                        elif ntxt != '':
                            break

                if not already_has:
                    body.insert(insert_after + 1, create_signature_paragraph(sig_label))
                modified_sigs += 1

        # 3. 每個新附表抬頭（例如「四 公共危險物品室內儲存場所抽查紀錄表」、「附表三...」、「附表十二...」）強制從新分頁開始
        TITLE_HINTS = ('抽查紀錄表', '數量一覽表', '種類及數量', '場所別名') + SUBTABLE_KEYWORDS
        cur_children = list(body)
        sub_table_seen = 0
        for idx, ch in enumerate(cur_children):
            if ch.tag != f"{W_NS}tbl" or ch is att10_tbl_ref:
                continue
            sub_table_seen += 1
            # 往前找該表格的抬頭段落（跳過最多 2 個空白段落）
            title_idx = -1
            for back in range(idx - 1, max(-1, idx - 4), -1):
                prev = cur_children[back]
                if prev.tag != f"{W_NS}p":
                    break
                ptxt = get_text(prev).strip()
                if ptxt == '':
                    continue
                if (
                    any(h in ptxt for h in TITLE_HINTS)
                    and not ptxt.startswith(('※', '一、', '二、', '備註', '附件10'))
                    and '保安檢查員簽章' not in ptxt
                    and '專人簽章' not in ptxt
                    and '本表欄位所稱辦法' not in ptxt
                ):
                    title_idx = back
                break

            if title_idx > 0:
                # 若是緊接在附件10分頁後的第一個附表，附件10結尾已有分頁符號，不重複加 pageBreakBefore
                if inserted_att10 and sub_table_seen == 1:
                    continue
                title_p = cur_children[title_idx]
                pPr = title_p.find('./w:pPr', NS)
                if pPr is None:
                    pPr = ET.Element(f"{W_NS}pPr")
                    title_p.insert(0, pPr)
                if pPr.find('./w:pageBreakBefore', NS) is None:
                    ET.SubElement(pPr, f"{W_NS}pageBreakBefore")
                # 清除抬頭段落正上方多餘的空白或手動分頁段落，避免出現雙重空白頁
                for back in range(title_idx - 1, max(-1, title_idx - 4), -1):
                    prev = cur_children[back]
                    if prev.tag == f"{W_NS}p" and get_text(prev).strip() == '':
                        if prev in list(body):
                            body.remove(prev)
                    else:
                        break

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
        'mode': '30倍以上' if is_above else '30倍以下',
        'inserted_att10': inserted_att10,
        'signatures': modified_sigs,
        'site_name': info['SITE_NAME'],
    }


def process_targets(targets, output_folder=None, mode='auto'):
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

    print("=" * 72)
    print(f"  公共危險物品場所自主檢查表 - 自動轉換程式 (v3.0 · 11308範例版)")
    print(f"  待處理檔案數量：{len(files_to_process)} 個  |  模式：{mode}")
    print("=" * 72)

    success_count = 0
    total_sigs = 0
    total_att10 = 0

    for i, src_file in enumerate(files_to_process, 1):
        file_dir = os.path.dirname(os.path.abspath(src_file))
        file_name = os.path.basename(src_file)

        if output_folder:
            dest_dir = os.path.abspath(output_folder)
        else:
            dest_dir = os.path.join(file_dir, "轉換完成")

        dst_file = os.path.join(dest_dir, file_name)

        try:
            ok, stats = convert_docx_file(src_file, dst_file, mode=mode)
            if ok:
                success_count += 1
                total_sigs += stats['signatures']
                if stats['inserted_att10']:
                    total_att10 += 1
                att_tag = "換附件10" if stats['inserted_att10'] else "僅轉附表"
                print(f"[{i:02d}/{len(files_to_process):02d}] 成功 ({stats['mode']} / {att_tag} / 附表簽章{stats['signatures']:02d}處): {file_name}")
            else:
                print(f"[{i:02d}/{len(files_to_process):02d}] 失敗: {file_name}")
        except Exception as e:
            print(f"[{i:02d}/{len(files_to_process):02d}] 發生錯誤: {file_name} -> {e}")

    print("=" * 72)
    print(f"  處理完畢！成功處理 {success_count} 個檔案：")
    print(f"  - 替換插入附件10主表：{total_att10} 份")
    print(f"  - 附表簽章改為範例樣式：{total_sigs} 處")
    if output_folder:
        print(f"  -> {os.path.abspath(output_folder)}")
    print("=" * 72)


if __name__ == '__main__':
    args = sys.argv[1:]
    out_folder = None
    mode = 'auto'
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
        elif a in ('-m', '--mode'):
            if i + 1 < len(args):
                mode = args[i + 1]
                skip_next = True
        else:
            filtered_args.append(a)

    if not filtered_args:
        default_input = r"路竹30倍以上-16家ok"
        default_output = r"路竹30倍以上-轉換完成"
        if os.path.exists(default_input):
            process_targets([default_input], output_folder=out_folder or default_output, mode=mode)
        else:
            cur_docx = glob.glob("*.docx")
            if cur_docx:
                process_targets(cur_docx, output_folder=out_folder or "轉換完成", mode=mode)
            else:
                print("請將檔案或資料夾拖曳至本程式上，或傳入路徑參數。")
    else:
        process_targets(filtered_args, output_folder=out_folder, mode=mode)

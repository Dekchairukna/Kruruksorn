"""Report form templates for Kruruksorn — fill official .docx forms in-web.

Each template is the teacher's original .docx (report_templates/<key>.docx);
generation only replaces the mapped field literals, so an unchanged field
yields byte-identical text and the layout/fonts stay exactly the same.
"""
import os
import docx

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE, "report_templates")


def _iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    def walk(tables):
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
                    for pp in walk(cell.tables):
                        yield pp
    for p in walk(doc.tables):
        yield p
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            if hf is None:
                continue
            for p in hf.paragraphs:
                yield p
            for t in hf.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p


def fill(src_path, out_path, mapping):
    """mapping: list of (old_literal, new_value). Longer literals first."""
    doc = docx.Document(src_path)
    for old, new in sorted(mapping, key=lambda x: -len(x[0])):
        if old == new or not old:
            continue
        for p in _iter_paragraphs(doc):
            if old in p.text and p.runs:
                full = "".join(r.text for r in p.runs)
                if old in full:
                    p.runs[0].text = full.replace(old, new)
                    for r in p.runs[1:]:
                        r.text = ""
    doc.save(out_path)


def template_path(key):
    return os.path.join(TEMPLATE_DIR, key + ".docx")


FORMS = {
    "analysis_math": {
        "title": "รายงานการวิเคราะห์ข้อสอบ (คณิตศาสตร์)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก  2569/",
                "find": [
                    "ที่  บก  2569/"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 3 เมษายน 2569",
                "find": [
                    "วันที่ 3 เมษายน 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ ๑",
                "find": [
                    "ภาคเรียนที่ ๑"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา ๒๕๖๙",
                "find": [
                    "ปีการศึกษา ๒๕๖๙"
                ]
            },
            {
                "key": "code",
                "label": "รหัสวิชา",
                "default": "ค33201",
                "find": [
                    "ค33201"
                ]
            }
        ]
    },
    "analysis_thai": {
        "title": "รายงานการวิเคราะห์ข้อสอบ (ภาษาไทย)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายปิยะทัศน์ แสงสว่าง",
                "profile": "full_name",
                "find": [
                    "นายปิยะทัศน์  แสงสว่าง",
                    "นายปิยะทัศน์ แสงสว่าง"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  วก  ๒๕๖๘ / ๑๕๕",
                "find": [
                    "ที่  วก  ๒๕๖๘ / ๑๕๕"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่  ๖   เดือน  ตุลาคม  พ.ศ. ๒๕๖๘",
                "find": [
                    "วันที่  ๖   เดือน  ตุลาคม  พ.ศ. ๒๕๖๘"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ ๑",
                "find": [
                    "ภาคเรียนที่ ๑"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา ๒๕๖๘",
                "find": [
                    "ปีการศึกษา ๒๕๖๘"
                ]
            },
            {
                "key": "code",
                "label": "รหัสวิชา",
                "default": "ท๒๓๑๐๑",
                "find": [
                    "ท๒๓๑๐๑"
                ]
            }
        ]
    },
    "best_practice": {
        "title": "รายงานวิธีการปฏิบัติที่เป็นเลิศ (Best Practice)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก 2569/202",
                "find": [
                    "ที่  บก 2569/202"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 16 เดือน กันยายน พ.ศ. 2569",
                "find": [
                    "วันที่ 16 เดือน กันยายน พ.ศ. 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา 2569",
                "find": [
                    "ปีการศึกษา 2569"
                ]
            },
            {
                "key": "code",
                "label": "รหัสวิชา",
                "default": "ค23101",
                "find": [
                    "ค23101"
                ]
            }
        ]
    },
    "grade_ptho5": {
        "title": "แบบบันทึกผลการเรียนประจำวิชา (ปถ.05)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก  ๒๕๖๙/",
                "find": [
                    "ที่  บก  ๒๕๖๙/"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 3 เดือน เมษายน พ.ศ. ๒๕๖๙",
                "find": [
                    "วันที่ 3 เดือน เมษายน พ.ศ. ๒๕๖๙"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา ๒๕๖9",
                "find": [
                    "ปีการศึกษา ๒๕๖9"
                ]
            }
        ]
    },
    "media_report": {
        "title": "แบบรายงานการใช้สื่อการเรียนการสอน",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก 2569/199",
                "find": [
                    "ที่  บก 2569/199"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 16 เดือน กันยายน พ.ศ. 2569",
                "find": [
                    "วันที่ 16 เดือน กันยายน พ.ศ. 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา 2569",
                "find": [
                    "ปีการศึกษา 2569"
                ]
            },
            {
                "key": "code",
                "label": "รหัสวิชา",
                "default": "ค23101",
                "find": [
                    "ค23101"
                ]
            }
        ]
    },
    "plc_plan": {
        "title": "แบบวิเคราะห์เพื่อออกแบบแผนการเรียนรู้ (PLC-PLAN)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก  2569/201",
                "find": [
                    "ที่  บก  2569/201"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 16 เดือน กันยายน พ.ศ. 2569",
                "find": [
                    "วันที่ 16 เดือน กันยายน พ.ศ. 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา 2569",
                "find": [
                    "ปีการศึกษา 2569"
                ]
            }
        ]
    },
    "plc_report": {
        "title": "รายงานผลกิจกรรม PLC",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่ บก 2569/227",
                "find": [
                    "ที่ บก 2569/227"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 23 กันยายน 2569",
                "find": [
                    "วันที่ 23 กันยายน 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา 2569",
                "find": [
                    "ปีการศึกษา 2569"
                ]
            }
        ]
    },
    "research_stad": {
        "title": "รายงานการวิจัยในชั้นเรียน (STAD)",
        "fields": [
            {
                "key": "teacher",
                "label": "ครูผู้สอน",
                "default": "นายสัญชัย แสนโคก",
                "profile": "full_name",
                "find": [
                    "นายสัญชัย  แสนโคก",
                    "นายสัญชัย แสนโคก"
                ]
            },
            {
                "key": "docno",
                "label": "เลขที่หนังสือ",
                "default": "ที่  บก 2569/200",
                "find": [
                    "ที่  บก 2569/200"
                ]
            },
            {
                "key": "date",
                "label": "วันที่",
                "default": "วันที่ 16 เดือน กันยายน พ.ศ. 2569",
                "find": [
                    "วันที่ 16 เดือน กันยายน พ.ศ. 2569"
                ]
            },
            {
                "key": "term",
                "label": "ภาคเรียน",
                "default": "ภาคเรียนที่ 1",
                "find": [
                    "ภาคเรียนที่ 1"
                ]
            },
            {
                "key": "year",
                "label": "ปีการศึกษา",
                "default": "ปีการศึกษา 2569",
                "find": [
                    "ปีการศึกษา 2569"
                ]
            },
            {
                "key": "code",
                "label": "รหัสวิชา",
                "default": "ค23101",
                "find": [
                    "ค23101"
                ]
            }
        ]
    }
}

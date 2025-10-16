"""
HWPX 파일 직접 생성 모듈 - 완전한 구조
"""
import os
import zipfile
from datetime import datetime
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
import logging

logger = logging.getLogger(__name__)


class HWPXWriter:
    """HWPX 파일 작성 클래스"""

    def __init__(self):
        pass

    def convert_docx_to_hwpx(self, docx_path: str, hwpx_path: str) -> bool:
        """
        DOCX 파일을 HWPX 파일로 변환
        """
        try:
            logger.info(f"변환 시작: {docx_path} -> {hwpx_path}")

            # DOCX 파일 로드
            doc = Document(docx_path)
            logger.info(f"DOCX 파일 로드 완료: {len(doc.paragraphs)}개 문단, {len(doc.tables)}개 표")

            # HWPX ZIP 파일 생성
            with zipfile.ZipFile(hwpx_path, 'w', zipfile.ZIP_DEFLATED) as hwpx:
                # 1. mimetype (압축 없이)
                hwpx.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)

                # 2. version.xml
                version_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"
    tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1"
    buildNumber="0" os="1" xmlVersion="1.5"
    application="Hancom Office Hangul" appVersion="12, 0, 0, 535 WIN32LEWindows_10"/>'''
                hwpx.writestr('version.xml', version_xml.encode('utf-8'))

                # 3. settings.xml
                settings_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">
    <ha:CaretPosition listIDRef="10" paraIDRef="1" pos="0"/>
</ha:HWPApplicationSetting>'''
                hwpx.writestr('settings.xml', settings_xml.encode('utf-8'))

                # 4. Contents/header.xml (간소화된 버전)
                header_xml = self._create_minimal_header_xml()
                hwpx.writestr('Contents/header.xml', header_xml.encode('utf-8'))

                # 5. Contents/section0.xml
                section_xml = self._create_section_xml(doc)
                hwpx.writestr('Contents/section0.xml', section_xml.encode('utf-8'))

                # 6. Contents/content.hpf (핵심!)
                content_hpf = self._create_content_hpf_xml()
                hwpx.writestr('Contents/content.hpf', content_hpf.encode('utf-8'))

                # 7. Scripts
                scripts_header = 'var Documents = XHwpDocuments;\nvar Document = Documents.Active_XHwpDocument;\n'
                scripts_source = 'function OnDocument_New()\n{\n\t//todo : \n}\n'
                hwpx.writestr('Scripts/headerScripts.js', scripts_header.encode('utf-16le'))
                hwpx.writestr('Scripts/sourceScripts.js', scripts_source.encode('utf-16le'))

                # 8. Preview/PrvText.txt
                preview_text = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
                hwpx.writestr('Preview/PrvText.txt', preview_text.encode('utf-8', errors='ignore'))

                # 9. META-INF/manifest.xml
                manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest>
    <file-entry full-path="version.xml" media-type="application/vnd.hancom.hwpml"/>
    <file-entry full-path="settings.xml" media-type="application/vnd.hancom.hwpml"/>
    <file-entry full-path="Contents/header.xml" media-type="application/vnd.hancom.hwpml"/>
    <file-entry full-path="Contents/section0.xml" media-type="application/vnd.hancom.hwpml"/>
    <file-entry full-path="Contents/content.hpf" media-type="application/vnd.hancom.hwpml"/>
</manifest>'''
                hwpx.writestr('META-INF/manifest.xml', manifest_xml.encode('utf-8'))

                # 10. META-INF/container.xml
                container_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0">
    <rootfiles>
        <rootfile full-path="Contents/content.hpf"/>
    </rootfiles>
</container>'''
                hwpx.writestr('META-INF/container.xml', container_xml.encode('utf-8'))

            logger.info(f"변환 완료: {hwpx_path}")
            return True

        except Exception as e:
            logger.error(f"변환 실패: {str(e)}")
            import traceback
            logger.error(f"상세 에러: {traceback.format_exc()}")
            return False

    def _create_minimal_header_xml(self):
        """최소한의 header.xml 생성"""
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
    xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
    version="1.5" secCnt="1">
    <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>
    <hh:refList>
        <hh:fontfaces itemCnt="1">
            <hh:fontface lang="HANGUL" fontCnt="1">
                <hh:font id="0" face="맑은 고딕" type="TTF" isEmbedded="0"/>
            </hh:fontface>
        </hh:fontfaces>
        <hh:borderFills itemCnt="2">
            <hh:borderFill id="1" threeD="0" shadow="0"/>
            <hh:borderFill id="2" threeD="0" shadow="0"/>
        </hh:borderFills>
        <hh:charProperties itemCnt="1">
            <hh:charPr id="0" height="1000" textColor="#000000" borderFillIDRef="2">
                <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
            </hh:charPr>
        </hh:charProperties>
        <hh:tabProperties itemCnt="1">
            <hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>
        </hh:tabProperties>
        <hh:paraProperties itemCnt="1">
            <hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0">
                <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
                <hh:heading type="NONE" idRef="0" level="0"/>
                <hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD"/>
                <hh:margin>
                    <hc:intent value="0" unit="HWPUNIT"/>
                    <hc:left value="0" unit="HWPUNIT"/>
                    <hc:right value="0" unit="HWPUNIT"/>
                </hh:margin>
                <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
                <hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"/>
            </hh:paraPr>
        </hh:paraProperties>
        <hh:styles itemCnt="1">
            <hh:style id="0" type="PARA" name="바탕글" engName="Normal"
                paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042"/>
        </hh:styles>
    </hh:refList>
</hh:head>'''

    def _create_section_xml(self, doc):
        """section0.xml 생성"""
        paragraphs_xml = []

        # 첫 문단에 섹션 정보 추가
        first_para_added = False

        for element in doc.element.body:
            if isinstance(element, CT_P):
                for para in doc.paragraphs:
                    if para._element == element:
                        if not first_para_added:
                            # 첫 문단에 섹션 설정 포함
                            para_xml = f'''<hp:p id="0" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
        <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000">
            <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>
            <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>
            <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0"/>
            <hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">
                <hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668" bottom="4252"/>
            </hp:pagePr>
            <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" fillArea="PAPER"/>
        </hp:secPr>
        <hp:t>{self._escape_xml(para.text)}</hp:t>
    </hp:run>
</hp:p>'''
                            first_para_added = True
                        else:
                            para_xml = f'''<hp:p id="0" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
        <hp:t>{self._escape_xml(para.text)}</hp:t>
    </hp:run>
</hp:p>'''
                        paragraphs_xml.append(para_xml)
                        break

            elif isinstance(element, CT_Tbl):
                for table in doc.tables:
                    if table._element == element:
                        table_xml = self._create_table_xml(table)
                        paragraphs_xml.append(table_xml)
                        break

        section_content = '\n'.join(paragraphs_xml)

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
    xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
    xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
{section_content}
</hs:sec>'''

    def _create_table_xml(self, table):
        """표를 XML로 변환"""
        rows_xml = []

        for row in table.rows:
            cells_xml = []
            for cell in row.cells:
                cell_text = self._escape_xml(cell.text)
                cell_xml = f'''<hp:tc>
        <hp:subList>
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0">
                <hp:run charPrIDRef="0">
                    <hp:t>{cell_text}</hp:t>
                </hp:run>
            </hp:p>
        </hp:subList>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
    </hp:tc>'''
                cells_xml.append(cell_xml)

            row_xml = f'''<hp:tr>
    {chr(10).join(cells_xml)}
</hp:tr>'''
            rows_xml.append(row_xml)

        return f'''<hp:p id="0" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
        <hp:tbl id="0" rowCnt="{len(table.rows)}" colCnt="{len(table.columns)}">
            {chr(10).join(rows_xml)}
        </hp:tbl>
    </hp:run>
</hp:p>'''

    def _create_content_hpf_xml(self):
        """content.hpf 생성 (핵심 파일)"""
        now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<opf:package xmlns:opf="http://www.idpf.org/2007/opf/"
    xmlns:dc="http://purl.org/dc/elements/1.1/" version="" unique-identifier="" id="">
    <opf:metadata>
        <opf:title>Converted Document</opf:title>
        <opf:language>ko</opf:language>
        <opf:meta name="CreatedDate" content="text">{now}</opf:meta>
        <opf:meta name="ModifiedDate" content="text">{now}</opf:meta>
    </opf:metadata>
    <opf:manifest>
        <opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>
        <opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>
        <opf:item id="headersc" href="Scripts/headerScripts.js" media-type="application/x-javascript ;charset=utf-16"/>
        <opf:item id="sourcesc" href="Scripts/sourceScripts.js" media-type="application/x-javascript ;charset=utf-16"/>
        <opf:item id="settings" href="settings.xml" media-type="application/xml"/>
    </opf:manifest>
    <opf:spine>
        <opf:itemref idref="header" linear="yes"/>
        <opf:itemref idref="section0" linear="yes"/>
        <opf:itemref idref="headersc" linear="yes"/>
        <opf:itemref idref="sourcesc" linear="yes"/>
    </opf:spine>
</opf:package>'''

    def _escape_xml(self, text):
        """XML 특수 문자 이스케이프"""
        if not text:
            return ""
        return (text.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace('"', "&quot;")
                    .replace("'", "&apos;"))

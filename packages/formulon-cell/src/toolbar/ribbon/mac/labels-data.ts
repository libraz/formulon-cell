import {
  type MacRibbonReasonStrings,
  macRibbonReasonsEn,
} from '../../../i18n/strings/mac/ribbon-reasons.js';
import type { Locale, Strings } from '../../../i18n/strings.js';
import { dictionaries } from '../../../i18n/strings.js';
import { MAC_BUNDLE_LABELS } from './labels-bundle.js';

export const MAC_LOCALES = ['ja', 'en', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export type MacLocale = (typeof MAC_LOCALES)[number];
export type Label = Readonly<Record<Locale, string>>;
export type LegacyLabel = Pick<Label, 'ja' | 'en'>;
export type LabelLike = Label | LegacyLabel;
export type ExtraTranslations = Readonly<{
  de: string;
  fr: string;
  zh: string;
  ko: string;
  th: string;
}>;

type RibbonKey = Exclude<keyof Strings['ribbon'], 'tabs' | 'fontFamilies'>;

/**
 * Mac labels without a reusable central dictionary field. Values are kept by
 * stable id and supplied for every non-English locale explicitly. Product
 * names and standards such as SmartArt, Python, PDF, and 100% intentionally
 * remain unchanged where that is the Excel UI term.
 */
export const MAC_LABEL_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  chartBar: { de: 'Balken', fr: 'Barres', zh: '条形图', ko: '가로 막대형', th: 'แท่ง' },
  chartArea: { de: 'Fläche', fr: 'Aires', zh: '面积图', ko: '영역형', th: 'พื้นที่' },
  shapeLine: { de: 'Linie', fr: 'Ligne', zh: '直线', ko: '선', th: 'เส้น' },
  shapeRectangle: { de: 'Rechteck', fr: 'Rectangle', zh: '矩形', ko: '직사각형', th: 'สี่เหลี่ยมผืนผ้า' },
  shapeOval: { de: 'Oval', fr: 'Ovale', zh: '椭圆', ko: '타원', th: 'วงรี' },
  shapeDiamond: { de: 'Raute', fr: 'Losange', zh: '菱形', ko: '마름모', th: 'ข้าวหลามตัด' },
  shapeArrow: { de: 'Pfeil', fr: 'Flèche', zh: '箭头', ko: '화살표', th: 'ลูกศร' },
  headerFooter: {
    de: 'Kopf- und Fußzeile',
    fr: 'En-tête et pied de page',
    zh: '页眉和页脚',
    ko: '머리글 및 바닥글',
    th: 'หัวกระดาษและท้ายกระดาษ',
  },
  penBlack: {
    de: 'Schwarzer Stift',
    fr: 'Stylet noir',
    zh: '黑色笔',
    ko: '검은색 펜',
    th: 'ปากกาสีดำ',
  },
  penRed: { de: 'Roter Stift', fr: 'Stylet rouge', zh: '红色笔', ko: '빨간색 펜', th: 'ปากกาสีแดง' },
  trackpad: {
    de: 'Mit Trackpad zeichnen',
    fr: 'Dessiner avec le Trackpad',
    zh: '使用触控板绘图',
    ko: '트랙피드로 그리기',
    th: 'วาดด้วยแทร็คแพ็ด',
  },
  pageSetup: {
    de: 'Seiteneinrichtung',
    fr: 'Mise en page',
    zh: '页面设置',
    ko: '페이지 설정',
    th: 'ตั้งค่าหน้ากระดาษ',
  },
  headingsHeading: {
    de: 'Überschriften',
    fr: 'Titres',
    zh: '标题',
    ko: '머리글',
    th: 'หัวเรื่อง',
  },
  dynamicArray: {
    de: 'Dynamisches Array',
    fr: 'Tableau dynamique',
    zh: '动态数组',
    ko: '동적 배열',
    th: 'อาร์เรย์แบบไดนามิก',
  },
  reset: { de: 'Zurücksetzen', fr: 'Réinitialiser', zh: '重置', ko: '다시 설정', th: 'รีเซ็ต' },
  textToColumns: {
    de: 'Text in Spalten',
    fr: 'Convertir',
    zh: '分列',
    ko: '텍스트 나누기',
    th: 'ข้อความเป็นคอลัมน์',
  },
  analysis: {
    de: 'Analyse-Funktionen',
    fr: 'Utilitaire d’analyse',
    zh: '分析工具库',
    ko: '분석 도구',
    th: 'เครื่องมือวิเคราะห์',
  },
  macros: {
    de: 'Makros anzeigen',
    fr: 'Afficher les macros',
    zh: '查看宏',
    ko: '매크로 보기',
    th: 'ดูแมโคร',
  },
};

/** Group titles worded by a ribbon dictionary entry. */
export const MAC_GROUP_SOURCES: Readonly<Record<string, RibbonKey>> = {
  Options: 'options',
};

export const MAC_GROUP_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  'Data from Picture': {
    de: 'Daten aus Bild',
    fr: 'Données de l’image',
    zh: '来自图片的数据',
    ko: '그림의 데이터',
    th: 'ข้อมูลจากรูปภาพ',
  },
  'Protection Changes': {
    de: 'Änderungen',
    fr: 'Modifications',
    zh: '更改',
    ko: '변경 내용',
    th: 'การเปลี่ยนแปลง',
  },
  'Queries & Connections': {
    de: 'Abfragen und Verbindungen',
    fr: 'Requêtes et connexions',
    zh: '查询和连接',
    ko: '쿼리 및 연결',
    th: 'คิวรีและการเชื่อมต่อ',
  },
};

type MacRibbonReason = keyof MacRibbonReasonStrings;

/** Why a ribbon command or menu item is unavailable, in every locale, from the dictionaries. */
export const MAC_CONTEXT_LABELS = Object.fromEntries(
  (Object.keys(macRibbonReasonsEn) as MacRibbonReason[]).map((key) => [
    key,
    Object.fromEntries(
      MAC_LOCALES.map((locale) => [locale, dictionaries[locale].mac.ribbonReasons[key]]),
    ),
  ]),
) as Readonly<Record<MacRibbonReason, Label>>;

export const MAC_MENU_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  'mac.page.theme.enableUpdates': {
    de: 'Inhaltsaktualisierungen aktivieren',
    fr: 'Activer les mises à jour de contenu',
    zh: '启用内容更新',
    ko: '콘텐츠 업데이트 사용',
    th: 'เปิดใช้งานการอัปเดตเนื้อหา',
  },
  'mac.page.theme.browse': {
    de: 'Nach Designs suchen...',
    fr: 'Rechercher des thèmes...',
    zh: '浏览主题...',
    ko: '테마 찾아보기...',
    th: 'เรียกดูธีม...',
  },
  'mac.page.theme.save': {
    de: 'Aktuelles Design speichern...',
    fr: 'Enregistrer le thème actif...',
    zh: '保存当前主题...',
    ko: '현재 테마 저장...',
    th: 'บันทึกธีมปัจจุบัน...',
  },
  // The View tab button; the Page Layout tab keeps the shared `pageLayout` label.
  'mac.view.pageLayout': {
    de: 'Seitenlayout',
    fr: 'Mode Page',
    zh: '页面布局',
    ko: '페이지 레이아웃',
    th: 'เค้าโครงหน้ากระดาษ',
  },
  'mac.automate.newScript.editor': {
    de: 'Im Code-Editor erstellen',
    fr: 'Créer dans l’Éditeur de code',
    zh: '在代码编辑器中创建',
    ko: '코드 편집기에서 만들기',
    th: 'สร้างในตัวแก้ไขโค้ด',
  },
  'mac.automate.newScript.recording': {
    de: 'Aus der Aufzeichnung erstellen',
    fr: 'Créer à partir de l’enregistrement',
    zh: '从录制内容创建',
    ko: '녹음/녹화에서 만들기',
    th: 'สร้างจากการบันทึก',
  },
  'mac.automate.showScripts.recent': {
    de: 'Zuletzt verwendete Skripte',
    fr: 'Scripts récents',
    zh: '最近使用的脚本',
    ko: '최근 스크립트',
    th: 'สคริปต์ล่าสุด',
  },
  'mac.automate.showScripts.workbook': {
    de: 'Skripte aus dieser Arbeitsmappe',
    fr: 'Scripts de ce classeur',
    zh: '此工作簿中的脚本',
    ko: '이 통합 문서의 스크립트',
    th: 'สคริปต์จากเวิร์กบุ๊กนี้',
  },
  'mac.automate.showScripts.samples': {
    de: 'Beispiele',
    fr: 'Exemples',
    zh: '示例',
    ko: '샘플',
    th: 'ตัวอย่าง',
  },
  'mac.review.accessibility.check': {
    de: 'Barrierefreiheit überprüfen',
    fr: 'Vérification de l’accessibilité',
    zh: '检查辅助功能',
    ko: '접근성 검사',
    th: 'ตรวจสอบการช่วยสำหรับการเข้าถึง',
  },
  'mac.review.accessibility.altText': {
    de: 'Alt Text',
    fr: 'Texte Alt',
    zh: '替换文本',
    ko: '대체 텍스트',
    th: 'ข้อความแสดงแทน',
  },
  'mac.review.accessibility.formatTable': {
    de: 'Als Tabelle formatieren',
    fr: 'Mettre sous forme de tableau',
    zh: '套用表格格式',
    ko: '표 서식',
    th: 'จัดรูปแบบเป็นตาราง',
  },
  'mac.review.accessibility.unmerge': {
    de: 'Zellverbund aufheben',
    fr: 'Annuler la fusion des cellules',
    zh: '取消单元格合并',
    ko: '셀 분할',
    th: 'ยกเลิกการผสานเซลล์',
  },
  'mac.review.accessibility.help': {
    de: 'Barrierefreiheit  Hilfe',
    fr: 'Accessibilité aide',
    zh: '辅助功能帮助',
    ko: '접근성 도움말',
    th: 'การช่วยสำหรับการเข้าถึงวิธีใช้',
  },
  'mac.review.hideInk.hide': {
    de: 'Freihand ausblenden',
    fr: 'Masquer les entrées manuscrites',
    zh: '隐藏墨迹',
    ko: '잉크 숨기기',
    th: 'ซ่อนหมึก',
  },
  'mac.review.hideInk.deleteSheet': {
    de: 'Alle Freihandelemente auf dem Blatt löschen',
    fr: 'Supprimer toutes les entrées manuscrites sur la feuille',
    zh: '删除工作表中的所有墨迹',
    ko: '시트의 모든 잉크 삭제',
    th: 'ลบหมึกทั้งหมดบนแผ่นงาน',
  },
  'mac.review.hideInk.deleteWorkbook': {
    de: 'Alle Freihandanmerkungen in der Arbeitsmappe löschen',
    fr: 'Supprimer tous les éléments manuscrits dans le classeur',
    zh: '删除工作簿中的所有墨迹',
    ko: '통합 문서에서 모든 잉크 삭제',
    th: 'ลบหมึกทั้งหมดในเวิร์กบุ๊ก',
  },
  // The Scale to Fit combos' last entry; the shared `pageScale.custom` keeps its own wording.
  'mac.review.previousComment': {
    de: 'Vorheriger Kommentar',
    fr: 'Commentaire précédent',
    zh: '上一条批注',
    ko: '이전 메모',
    th: 'ข้อคิดเห็นก่อนหน้า',
  },
  'mac.review.nextComment': {
    de: 'Nächster Kommentar',
    fr: 'Commentaire suivant',
    zh: '下一条批注',
    ko: '다음 메모',
    th: 'ข้อคิดเห็นถัดไป',
  },
  'mac.formulas.removeArrows.all': {
    de: 'Pfeile entfernen',
    fr: 'Supprimer les flèches',
    zh: '删除箭头',
    ko: '연결선 제거',
    th: 'เอาลูกศรออก',
  },
  'mac.formulas.reset.clear': {
    de: 'Zurücksetzen',
    fr: 'Réinitialiser',
    zh: '重置',
    ko: '다시 설정',
    th: 'รีเซ็ต',
  },
  'mac.formulas.reset.runtime': {
    de: 'Laufzeit zurücksetzen',
    fr: 'Réinitialiser le délaid’exécution',
    zh: '重置运行时',
    ko: '런타임 다시 설정',
    th: 'รีเซ็ตรันไทม์',
  },
  'mac.formulas.reset.versionStable': {
    de: 'Stabil',
    fr: 'Stable',
    zh: '稳定',
    ko: '안정적',
    th: 'เสถียร',
  },
  'mac.formulas.reset.versionLatest': {
    de: 'Neueste',
    fr: 'Le plus récent',
    zh: '最新',
    ko: '최신',
    th: 'ล่าสุด',
  },
  'mac.formulas.reset.versionLearnMore': {
    de: 'Weitere Informationen',
    fr: 'En savoir plus',
    zh: '了解详细信息',
    ko: '자세히 알아보기',
    th: 'เรียนรู้เพิ่มเติม',
  },
  'mac.formulas.defineName.define': {
    de: 'Namen definieren...',
    fr: 'Définir un nom...',
    zh: '定义名称...',
    ko: '이름 정의...',
    th: 'กําหนดชื่อ...',
  },
  'mac.formulas.defineName.apply': {
    de: 'Übernehmen...',
    fr: 'Appliquer...',
    zh: '应用...',
    ko: '적용...',
    th: 'นําไปใช้...',
  },
  'mac.draw.penBlack': {
    de: 'Stift: Schwarz, 0.35 mm',
    fr: 'Stylo\u00a0: Noir, 0.35 mm',
    zh: '笔: 黑色，0.35 毫米',
    ko: '펜: 검정, 0.35mm',
    th: 'ปากกา: ดำ, 0.35 มม.',
  },
  'mac.draw.penRed': {
    de: 'Stift: Rot, 0.35 mm',
    fr: 'Stylo\u00a0: Rouge, 0.35 mm',
    zh: '笔: 红色，0.35 毫米',
    ko: '펜: 빨강, 0.35mm',
    th: 'ปากกา: แดง, 0.35 มม.',
  },
  'mac.draw.pencil': {
    de: 'Bleistift: Grau, 0.5 mm',
    fr: 'Crayon\u00a0: Gris, 0.5 mm',
    zh: '铅笔: 灰色，0.5 毫米',
    ko: '연필: 회색, 0.5mm',
    th: 'ดินสอ: เทา, 0.5 มม.',
  },
  'mac.draw.highlighter': {
    de: 'Textmarker: Gelb, 6 mm',
    fr: 'Surligneur\u00a0: Jaune, 6 mm',
    zh: '荧光笔: 黄色，6 毫米',
    ko: '형광펜: 노랑, 6mm',
    th: 'ปากกาเน้นข้อความ: เหลือง, 6 มม.',
  },
  'mac.draw.eraser.stroke': {
    de: 'Strichradierer',
    fr: 'Effaceur de traits',
    zh: '笔划橡皮擦',
    ko: '스트로크 지우개',
    th: 'ยางลบสำหรับเส้นรอบรูป',
  },
  'mac.draw.eraser.small': {
    de: 'Kleiner Radierer',
    fr: 'Petite gomme',
    zh: '小橡皮擦',
    ko: '작은 지우개',
    th: 'ยางลบขนาดเล็ก',
  },
  'mac.draw.eraser.medium': {
    de: 'Mittlerer Radierer',
    fr: 'Gomme moyenne',
    zh: '中橡皮擦',
    ko: '중간 지우개',
    th: 'ยางลบขนาดกลาง',
  },
  'mac.draw.eraser.segment': {
    de: 'Segmentradierer',
    fr: 'Effaceur de segments',
    zh: '线段橡皮擦',
    ko: '세그먼트 지우개',
    th: 'ยางลบเซกเมนต์',
  },
  'mac.draw.add.pen': {
    de: 'Stift',
    fr: 'Stylo',
    zh: '笔',
    ko: '펜',
    th: 'ปากกา',
  },
  'mac.draw.add.highlighter': {
    de: 'Textmarker',
    fr: 'Surligneur',
    zh: '荧光笔',
    ko: '형광펜',
    th: 'ปากกาเน้นข้อความ',
  },
  'mac.draw.add.pencil': {
    de: 'Bleistift',
    fr: 'Crayon',
    zh: '铅笔',
    ko: '연필',
    th: 'ดินสอ',
  },
  'mac.insert.forms.new': {
    de: 'Neues Formular',
    fr: 'Nouveau formulaire',
    zh: '新建表单',
    ko: '새 양식',
    th: 'ฟอร์มใหม่',
  },
  'mac.insert.forms.preview': {
    de: 'Vorschau von Formular anzeigen',
    fr: 'Aperçu du formulaire',
    zh: '预览表单',
    ko: '양식 미리 보기',
    th: 'แสดงตัวอย่างฟอร์ม',
  },
  'mac.insert.forms.edit': {
    de: 'Formular bearbeiten',
    fr: 'Modifier le formulaire',
    zh: '编辑表单',
    ko: '양식 편집',
    th: 'แก้ไขฟอร์ม',
  },
  'mac.insert.forms.send': {
    de: 'Formular senden',
    fr: 'Envoyer le formulaire',
    zh: '发送表单',
    ko: '양식 보내기',
    th: 'ส่งฟอร์ม',
  },
  'mac.insert.pictureFromData.file': {
    de: 'Bild aus Datei...',
    fr: 'Image à partir d’un fichier...',
    zh: '来自文件的图片…',
    ko: '파일에서 그림...',
    th: 'รูปภาพจากไฟล์...',
  },
  'mac.insert.pictureFromData.clipboard': {
    de: 'Bild aus Zwischenablage',
    fr: 'Image à partir du Presse-papiers',
    zh: '来自剪贴板的图片',
    ko: '클립보드의 사진',
    th: 'รูปภาพจากคลิปบอร์ด',
  },
  'mac.insert.photo.browser': {
    de: 'Fotobrowser…',
    fr: 'Navigateur de photos…',
    zh: '照片浏览器…',
    ko: '사진 브라우저…',
    th: 'เรียกดูรูปภาพ...',
  },
  'mac.insert.photo.file': {
    de: 'Bild aus Datei...',
    fr: 'Image à partir d’un fichier...',
    zh: '来自文件的图片...',
    ko: '파일에서 그림 삽입...',
    th: 'รูปภาพจากไฟล์...',
  },
  'mac.insert.photo.stock': {
    de: 'Archivbilder...',
    fr: 'Images de photothèque...',
    zh: '图像集…',
    ko: '스톡 이미지...',
    th: 'รูปในคลัง...',
  },
  'mac.insert.photo.online': {
    de: 'Onlinebilder...',
    fr: 'Images en ligne...',
    zh: '联机图片...',
    ko: '온라인 그림...',
    th: 'รูปภาพออนไลน์...',
  },
  'mac.insert.threeDModel.device': {
    de: 'Dieses Gerät...',
    fr: 'Cet appareil...',
    zh: '此设备...',
    ko: '이 디바이스...',
    th: 'อุปกรณ์นี้...',
  },
  'mac.insert.threeDModel.stock': {
    de: 'Archiv-3D-Modelle...',
    fr: 'les modèles 3D boursiers...',
    zh: '联机 3D 模型...',
    ko: '스톡 3D 모델...',
    th: 'โมเดล 3D ในสต็อก...`O3',
  },
  'mac.insert.smartArt.list': {
    de: 'Liste',
    fr: 'Liste',
    zh: '列表',
    ko: '목록',
    th: 'รายการ',
  },
  'mac.insert.smartArt.process': {
    de: 'Prozess',
    fr: 'Processus',
    zh: '流程',
    ko: '프로세스형',
    th: 'กระบวนการ',
  },
  'mac.insert.smartArt.cycle': {
    de: 'Zyklus',
    fr: 'Cycle',
    zh: '循环',
    ko: '주기형',
    th: 'วงกลม',
  },
  'mac.insert.smartArt.hierarchy': {
    de: 'Hierarchie',
    fr: 'Hiérarchie',
    zh: '层次结构',
    ko: '계층 구조형',
    th: 'ลำดับชั้น',
  },
  'mac.insert.smartArt.relationship': {
    de: 'Beziehung',
    fr: 'Relation',
    zh: '关系',
    ko: '관계형',
    th: 'ความสัมพันธ์',
  },
  'mac.insert.smartArt.matrix': {
    de: 'Matrix',
    fr: 'Matrice',
    zh: '矩阵',
    ko: '행렬형',
    th: 'เมทริกซ์',
  },
  'mac.insert.smartArt.pyramid': {
    de: 'Pyramide',
    fr: 'Pyramide',
    zh: '棱锥图',
    ko: '피라미드형',
    th: 'พีระมิด',
  },
  'mac.insert.smartArt.picture': {
    de: 'Bild',
    fr: 'Image',
    zh: '图片',
    ko: '그림',
    th: 'รูปภาพ',
  },
  'mac.insert.screenshot.clipping': {
    de: 'Bildschirmausschnitt',
    fr: 'Capture d’écran',
    zh: '屏幕剪辑',
    ko: '화면 캡처',
    th: 'การคลิปหน้าจอ',
  },
  'mac.insert.sparkline.line': {
    de: 'Linie',
    fr: 'Courbe',
    zh: '折线',
    ko: '꺾은선형',
    th: 'เส้น',
  },
  'mac.insert.sparkline.column': {
    de: 'Säule',
    fr: 'Colonne',
    zh: '柱形',
    ko: '세로 막대형',
    th: 'คอลัมน์',
  },
  'mac.insert.sparkline.winLoss': {
    de: 'Gewinn/Verlust',
    fr: 'Conclusions et pertes',
    zh: '盈亏',
    ko: '승패\u200b\u200b',
    th: 'ชนะ/แพ้',
  },
  'mac.insert.textBox.horizontal': {
    de: 'Horizontales Textfeld',
    fr: 'Zone de texte horizontale',
    zh: '横排文本框',
    ko: '가로 텍스트 상자',
    th: 'กล่องข้อความแนวนอน',
  },
  'mac.insert.textBox.vertical': {
    de: 'Vertikales Textfeld',
    fr: 'Zone de texte verticale',
    zh: '竖排文本框',
    ko: '세로 텍스트 상자',
    th: 'กล่องข้อความแนวตั้ง',
  },
  'mac.insert.equation.insert': {
    de: 'Formel',
    fr: 'Équation',
    zh: '公式',
    ko: '수식',
    th: 'สมการ',
  },

  'mac.data.validation.clearRules': {
    de: 'Datenüberprüfung löschen',
    fr: 'Effacer la validation des données',
    zh: '清除数据验证',
    ko: '데이터 유효성 검사 지우기',
    th: 'ล้างการตรวจสอบความถูกต้องของข้อมูล',
  },
  'mac.page.margins.normal': { de: 'Normal', fr: 'Normales', zh: '普通', ko: '보통', th: 'ปกติ' },
  'mac.page.margins.wide': { de: 'Breit', fr: 'Large', zh: '宽', ko: '넓게', th: 'กว้าง' },
  'mac.page.margins.narrow': { de: 'Schmal', fr: 'Étroites', zh: '窄', ko: '좁게', th: 'แคบ' },
  'mac.page.orientation.portrait': {
    de: 'Hochformat',
    fr: 'Portrait',
    zh: '纵向',
    ko: '세로',
    th: 'แนวตั้ง',
  },
  'mac.page.orientation.landscape': {
    de: 'Querformat',
    fr: 'Paysage',
    zh: '横向',
    ko: '가로',
    th: 'แนวนอน',
  },
  'mac.page.size.a4': { de: 'A4', fr: 'A4', zh: 'A4', ko: 'A4', th: 'A4' },
  'mac.page.size.a3': { de: 'A3', fr: 'A3', zh: 'A3', ko: 'A3', th: 'A3' },
  'mac.page.size.letter': {
    de: 'Letter',
    fr: 'Lettre',
    zh: '信纸',
    ko: 'Letter',
    th: 'Letter',
  },
  'mac.page.size.legal': {
    de: 'Legal',
    fr: 'Juridique',
    zh: '法律',
    ko: 'Legal',
    th: 'Legal',
  },
  'mac.page.printArea.add': {
    de: 'Zum Druckbereich hinzufügen',
    fr: 'Ajouter à la zone d’impression',
    zh: '添加到打印区域',
    ko: '인쇄 영역에 추가',
    th: 'เพิ่มลงในพื้นที่พิมพ์',
  },
  'mac.review.editNote': {
    de: 'Notiz bearbeiten',
    fr: 'Modifier la note',
    zh: '编辑备注',
    ko: '메모 편집',
    th: 'แก้ไขบันทึกย่อ',
  },
  'mac.view.freeze.panes': {
    de: 'Fenster fixieren',
    fr: 'Figer les volets',
    zh: '冻结窗格',
    ko: '틀 고정',
    th: 'ตรึงแนว',
  },
  'mac.view.freeze.off': {
    de: 'Fixierung aufheben',
    fr: 'Libérer les volets',
    zh: '取消冻结窗格',
    ko: '틀 고정 취소',
    th: 'ยกเลิกการตรึงแนว',
  },
};

const readCoreValue = (locale: Locale, key: RibbonKey): string => {
  const value = (dictionaries[locale].ribbon as unknown as Record<string, unknown>)[key];
  if (typeof value !== 'string')
    throw new Error(`Mac label source ribbon.${String(key)} is not a string.`);
  return value;
};

export const localizedFromCore = (key: RibbonKey): Label =>
  Object.fromEntries(MAC_LOCALES.map((locale) => [locale, readCoreValue(locale, key)])) as Label;

export const localizedFromTranslations = (
  seed: LegacyLabel,
  translations: ExtraTranslations,
): Label => ({ ...seed, ...translations });

export const localizedGroupLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`groups:${key}`];
  if (bundled) return bundled;
  const source = MAC_GROUP_SOURCES[key];
  if (source) return localizedFromCore(source);
  const translations = MAC_GROUP_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for group: ${key}`);
};

export const localizedMenuLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`menu:${key}`];
  if (bundled) return bundled;
  const translations = MAC_MENU_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for menu item: ${key}`);
};

/** Common function aliases used while a native live catalog is still being
 * assembled. A live engine's localized metadata supersedes these values. */
export const MAC_FUNCTION_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  SUM: { de: 'SUMME', fr: 'SOMME', zh: 'SUM', ko: 'SUM', th: 'SUM' },
  AVERAGE: { de: 'MITTELWERT', fr: 'MOYENNE', zh: 'AVERAGE', ko: 'AVERAGE', th: 'AVERAGE' },
  COUNT: { de: 'ANZAHL', fr: 'NB', zh: 'COUNT', ko: 'COUNT', th: 'COUNT' },
  COUNTA: { de: 'ANZAHL2', fr: 'NBVAL', zh: 'COUNTA', ko: 'COUNTA', th: 'COUNTA' },
  COUNTIF: { de: 'ZÄHLENWENN', fr: 'NB.SI', zh: 'COUNTIF', ko: 'COUNTIF', th: 'COUNTIF' },
  SUMIF: { de: 'SUMMEWENN', fr: 'SOMME.SI', zh: 'SUMIF', ko: 'SUMIF', th: 'SUMIF' },
  SUMIFS: { de: 'SUMMEWENNS', fr: 'SOMME.SI.ENS', zh: 'SUMIFS', ko: 'SUMIFS', th: 'SUMIFS' },
  IF: { de: 'WENN', fr: 'SI', zh: 'IF', ko: 'IF', th: 'IF' },
  AND: { de: 'UND', fr: 'ET', zh: 'AND', ko: 'AND', th: 'AND' },
  OR: { de: 'ODER', fr: 'OU', zh: 'OR', ko: 'OR', th: 'OR' },
  NOT: { de: 'NICHT', fr: 'NON', zh: 'NOT', ko: 'NOT', th: 'NOT' },
  XLOOKUP: { de: 'XVERWEIS', fr: 'RECHERCHEX', zh: 'XLOOKUP', ko: 'XLOOKUP', th: 'XLOOKUP' },
  VLOOKUP: { de: 'SVERWEIS', fr: 'RECHERCHEV', zh: 'VLOOKUP', ko: 'VLOOKUP', th: 'VLOOKUP' },
  HLOOKUP: { de: 'WVERWEIS', fr: 'RECHERCHEH', zh: 'HLOOKUP', ko: 'HLOOKUP', th: 'HLOOKUP' },
  MATCH: { de: 'VERGLEICH', fr: 'EQUIV', zh: 'MATCH', ko: 'MATCH', th: 'MATCH' },
  ROUND: { de: 'RUNDEN', fr: 'ARRONDI', zh: 'ROUND', ko: 'ROUND', th: 'ROUND' },
  TODAY: { de: 'HEUTE', fr: 'AUJOURDHUI', zh: 'TODAY', ko: 'TODAY', th: 'TODAY' },
  NOW: { de: 'JETZT', fr: 'MAINTENANT', zh: 'NOW', ko: 'NOW', th: 'NOW' },
  DATE: { de: 'DATUM', fr: 'DATE', zh: 'DATE', ko: 'DATE', th: 'DATE' },
};

export const localizedFunctionName = (name: string, locale: Locale): string => {
  if (locale === 'ja' || locale === 'en') return name;
  const translations = MAC_FUNCTION_TRANSLATIONS[name.toUpperCase()];
  return translations?.[locale] ?? name;
};

export const localizedLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`labels:${key}`];
  if (bundled) return bundled;
  const translations = MAC_LABEL_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for label: ${key}`);
};

package html2Word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;

public class CreateWord {

	public static void main(String[] args) throws Exception {

		// createWord();

		parseHtml();
	}

	public static void createWord() throws Exception {
		printGroupWord(null);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void parseHtml() throws Exception {
		Map map = new HashMap();
		File dic = new File("E:/work/1-hongye/滨州/数据库表结构/122ORCL");
		File[] files = dic.listFiles();
		for (File input : files) {
			FileInputStream is = new FileInputStream(input);
			byte[] b = new byte[is.available()];
			is.read(b);
			String str = new String(b, "GBK");
			Document doc = Jsoup.parse(str);

			Element title = doc.getElementsByClass("MAIN_TITLE").first();
			String tableName = title.html().replaceAll("Table", ""); // 表名
			String ptd = title.parent().html().replaceAll("<p.*</p>", "");
			ptd = ("".equals(ptd) ? tableName : ptd); // 中文描述
			String key = ptd.trim() + "<" + tableName.trim() + ">";
			System.err.println(key);

			Element table = doc.getElementsByClass("SIMPLE_TABLE").first();
			Elements trs = table.getElementsByTag("tr");
			List list = new ArrayList();
			for (int i = 1; i < trs.size(); i++) {
				Element tr = trs.get(i);
				Elements tds = tr.children();
				TableObj obj = new TableObj();
				String name = tds.get(0).html().replaceAll("</a>", "").replaceAll("<a href=.*>", "")
						.replaceAll("&nbsp;", "");
				obj.setName(name);
				String type = tds.get(1).html().replaceAll("&nbsp;", "");
				obj.setType(type);
				String optional = tds.get(2).html().replaceAll("&nbsp;", "");
				obj.setOptional(optional);
				String defaultValue = tds.get(3).html().replaceAll("&nbsp;", "");
				obj.setDefaultValue(defaultValue);
				String comments = tds.get(4).html().replaceAll("&nbsp;", "");
				obj.setComments(comments);
				obj.setIsPK(("ID".equals(obj.getName()) ? "是" : "否"));
				list.add(obj);
			}
			map.put(key, list);
			is.close();
		}
		printGroupWord(map);
	}

	@SuppressWarnings("rawtypes")
	public static void printGroupWord(Map map) throws Exception {
		XWPFDocument doc = new XWPFDocument();
		Set set = map.entrySet();
		Iterator itr = set.iterator();
		while (itr.hasNext()) {
			// List list = itr.next();
			Map.Entry entry = (Entry) itr.next();
			String key = (String) entry.getKey();
			List list = (List) entry.getValue();
			XWPFParagraph p1 = doc.createParagraph();
			int size = list.size();

			XWPFTable table = doc.createTable(size + 1, 6);
			// 设置上下左右四个方向的距离，可以将表格撑大
			// 表格属性
			CTTblPr tablePr = table.getCTTbl().addNewTblPr();
			// 表格宽度
			CTTblWidth width = tablePr.addNewTblW();
			width.setW(BigInteger.valueOf(8000));
			// table.set
			List<XWPFTableCell> tableCells1 = table.getRow(0).getTableCells();
			tableCells1.get(0).setText("代码");
			tableCells1.get(1).setText("名称");
			tableCells1.get(2).setText("数据类型");
			tableCells1.get(3).setText("是否主键");
			tableCells1.get(4).setText("是否必填");
			tableCells1.get(5).setText("注释");

			for (int i = 0; i < size; i++) {
				TableObj info = (TableObj) list.get(i);
				List<XWPFTableCell> tableCells = table.getRow(i + 1).getTableCells();
				tableCells.get(0).setText(info.getName());
				tableCells.get(1).setText(info.getComments());
				tableCells.get(2).setText(info.getType());
				tableCells.get(3).setText(info.getIsPK());
				tableCells.get(4).setText("");
				tableCells.get(5).setText(info.getComments());
			}

			// 设置字体对齐方式
			p1.setAlignment(ParagraphAlignment.CENTER);
			p1.setVerticalAlignment(TextAlignment.TOP);
			// 第一页要使用p1所定义的属性
			XWPFRun r1 = p1.createRun();
			// 设置字体是否加粗
			r1.setBold(true);
			r1.setFontSize(20);
			// 设置使用何种字体
			r1.setFontFamily("Courier");
			// 设置上下两行之间的间距
			r1.setTextPosition(20);
			r1.setText(key);
		}

		FileOutputStream out = new FileOutputStream("d://table_struct.docx");
		doc.write(out);
		out.close();
	}

}

class TableObj {
	private String name;

	private String isPK;

	public String getIsPK() {
		return isPK;
	}

	public void setIsPK(String isPK) {
		this.isPK = isPK;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public String getOptional() {
		return optional;
	}

	public void setOptional(String optional) {
		this.optional = optional;
	}

	public String getDefaultValue() {
		return defaultValue;
	}

	public void setDefaultValue(String defaultValue) {
		this.defaultValue = defaultValue;
	}

	public String getComments() {
		return comments;
	}

	public void setComments(String comments) {
		this.comments = comments;
	}

	private String type;
	private String optional;
	private String defaultValue;
	private String comments;
}
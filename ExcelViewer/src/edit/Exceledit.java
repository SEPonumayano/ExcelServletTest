package edit;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Servlet implementation class Exceledit
 */
public class Exceledit extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * @see HttpServlet#HttpServlet()
     */
    public Exceledit() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		InputStream is =null;
		Workbook wb=null;

		String name=null;
		String value=null;
		String number=null;
		String date=null;

		try {

			//読み込みたいファイル
			is=new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");
			wb=WorkbookFactory.create(is);

			//どこのシート？
			Sheet sh =wb.getSheet("sheet1");

			//name
			Row row1=sh.getRow(1);
			Cell cell1=row1.getCell(1);
			name=cell1.getStringCellValue();


			//value
			Row row3=sh.getRow(3);
			Cell cell2 =row3.getCell(2);
			value=cell2.getStringCellValue();


			//number
			Row row0=sh.getRow(0);
			Cell cell3=row0.getCell(3);
			double cellnumber=cell3.getNumericCellValue();
			number = String.valueOf((int)cellnumber);

			//date
			Row row7=sh.getRow(7);
			Cell cell0=row7.getCell(0);
			//ユーザー定義されているdateの値を表示するための変換
			if(BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX <=cell0.getCellStyle().getDataFormat()) {
				CellFormat cellFormat =CellFormat.getInstance(
						cell0.getCellStyle().getDataFormatString());
				CellFormatResult cellFormatResult=cellFormat.apply(cell0);
				date=cellFormatResult.text;
			}



		}catch(Exception ex){
			ex.printStackTrace();
		}


		request.setAttribute("name", name);
		request.setAttribute("value", value);
		request.setAttribute("number",number);
		request.setAttribute("date", date);

		RequestDispatcher rd = request.getRequestDispatcher("/edit.jsp");
		rd.forward(request, response);

	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub

		response.setContentType("text/html; charset=UTF-8");
    	request.setCharacterEncoding("UTF-8");

    	String name=request.getParameter("name");
    	String number=request.getParameter("number");
    	String date=request.getParameter("date");
    	String value=request.getParameter("value");

    	//数字にカンマ付けの制御する場合はデータ型の変換が必要
    	int num =Integer.parseInt(number);

    	//日付型への変換(※フォーマットの形で入力しないと登録できないのでバリテーション必須)
    	SimpleDateFormat day=new SimpleDateFormat("yyyy/MM/dd");
    	Date formatDate=null;
    	try {
			formatDate =day.parse(date);
		} catch (ParseException e1) {
			// TODO 自動生成された catch ブロック
			e1.printStackTrace();
		}



		InputStream is =null;
		Workbook wb=null;

		try {
			is=new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");
			wb=WorkbookFactory.create(is);

		}catch(IOException e) {
			System.out.println(e.toString());
		}finally {
			try {
				is.close();
			}catch(IOException e) {
				System.out.println(e.toString());
			}
		}

		//値の場所指定
			Sheet sh =wb.getSheet("sheet1");
			Row row5=sh.createRow(1);
			Cell cell5=row5.createCell(1);

			//↑の時点で使う部分のRowは一通り定義しておくと便利らしい…
			// ex) Row row0=sh.createRow(0); Row row1=sh.createRow(1);

			cell5.setCellValue(name);

		//スタイル編集
			//文字
			Font font=wb.createFont();
			//文字色
			font.setColor(IndexedColors.DARK_RED.index);
			CellStyle cellstyle=wb.createCellStyle();
			cellstyle.setFont(font);
			//セル
			cellstyle.setFillForegroundColor(IndexedColors.CORAL.index);
			cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell5.setCellStyle(cellstyle);

		//値の代入
			//cell.setCellValue(name);

		//セルの結合 ※結合設定は一回きりなので、新規ファイル作成以外には書かないほうが無難
			//CellRangeAddress cra=new CellRangeAddress(3,4,2,2);   // (int FirstRow,int LastRow,int FirstCell,int LastCell)
			//sh.addMergedRegion(cra);
			//Row row1=sh.createRow(cra.getFirstRow());
			//Cell cell1 =row1.createCell(cra.getFirstColumn());
			//cell1.setCellValue("結合しました");

			//結合したセルは位置指定すれば値の編集可能
			Row row1=sh.createRow(3);
			Cell cell1 =row1.createCell(2);
			cell1.setCellValue(value);


		//値を日付型で登録
			Row row2 =sh.createRow(7);
			Cell cell2=row2.createCell(0);

			//現在時刻を登録
			//cell2.setCellValue(Calendar.getInstance());
			//任意の時刻を登録
			//cell2.setCellValue("2020/09/09 7:08");
			//入力値を登録
			cell2.setCellValue(formatDate);

			CreationHelper createHelper=wb.getCreationHelper();
			CellStyle cs=wb.createCellStyle();
			short style=createHelper.createDataFormat().getFormat("yyyy/mm/dd");
			cs.setDataFormat(style);
			//日付型のスタイルを登録
			cell2.setCellStyle(cs);

		//数値にカンマをつける
			Row row3=sh.createRow(0);
			Cell cell3=row3.createCell(3);
			cell3.setCellValue(num);

			DataFormat fm =wb.createDataFormat();
			CellStyle cs2=wb.createCellStyle();
			cs2.setDataFormat(fm.getFormat("#,##0"));
			cell3.setCellStyle(cs2);


			FileOutputStream out =null;
			try {
				//ここに返します
				out=new FileOutputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");

				//編集部分を書いて保存しまーす
				wb.write(out);

			}catch(IOException e) {
				System.out.println(e.toString());
			}finally {
				try {
					out.close();
				}catch(IOException e) {
					System.out.println(e.toString());
				}
		}
			RequestDispatcher rd = request.getRequestDispatcher("/ok.jsp");
			rd.forward(request, response);


		}


}

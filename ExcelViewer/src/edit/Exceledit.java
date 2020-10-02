package edit;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub

		response.setContentType("text/html; charset=UTF-8");
    	request.setCharacterEncoding("UTF-8");

    	String name=request.getParameter("name");

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
			Row row=sh.createRow(0);
			Cell cell=row.createCell(0);

		//スタイル編集
			//文字
			Font font=wb.createFont();
			//文字色
			font.setColor(IndexedColors.DARK_RED.index);
			CellStyle cellstyle=wb.createCellStyle();
			cellstyle.setFont(font);

			//セル
			cellstyle.setFillForegroundColor(IndexedColors.LAVENDER.index);
			cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(cellstyle);

		//値の代入
			cell.setCellValue(name);

		//セルの結合 ※結合設定は一回きりなので、書かないほうが無難
			//CellRangeAddress cra=new CellRangeAddress(3,4,2,2);   // (int FirstRow,int LastRow,int FirstCell,int LastCell)
			//sh.addMergedRegion(cra);
			//Row row1=sh.createRow(cra.getFirstRow());
			//Cell cell1 =row1.createCell(cra.getFirstColumn());
			//cell1.setCellValue("結合しました");

			//結合したセルは位置指定すれば値の編集可能
			Row row1=sh.createRow(3);
			Cell cell1 =row1.createCell(2);
			cell1.setCellValue("追加");
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

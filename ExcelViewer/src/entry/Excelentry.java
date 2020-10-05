package entry;

import java.io.FileOutputStream;
import java.io.IOException;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




/**
 * Servlet implementation class Excelentry
 */
public class Excelentry extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * @see HttpServlet#HttpServlet()
     */
    public Excelentry() {
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
    	String price=request.getParameter("price");
    	String filename=request.getParameter("filename");

    	Workbook wb = new XSSFWorkbook();

		//①書き込みたいシート
		Sheet sheet1 = wb.createSheet();

		//②どこの行？※1列目=0
		Row row1=sheet1.createRow(0);

		//③どこの列？
		Cell cell1=row1.createCell(0);

		//④書き込みたいこと
		cell1.setCellValue(name);

		Row row2=sheet1.createRow(1);
		Cell cell2=row2.createCell(0);
		cell2.setCellValue(price);

		FileOutputStream out =null;

		try {
			//ここに返します
			out=new FileOutputStream("C:\\Users\\onumaayano1199\\Pictures\\"+filename+".xlsx");

			//編集部分を書いて保存しまーす
			wb.write(out);

		}catch(IOException e) {
			System.out.println(e.toString());
		}finally {
			try {
				wb.close();
				out.close();
			}catch(IOException e) {
				System.out.println(e.toString());
			}
	}
		RequestDispatcher rd = request.getRequestDispatcher("/ok.jsp");
		rd.forward(request, response);

}}

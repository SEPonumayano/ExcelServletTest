package input;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Servlet implementation class Excelinput
 */
public class Excelinput extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * @see HttpServlet#HttpServlet()
     */
    public Excelinput() {
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

		String value=null;
		String value1=null;

		try {

			//読み込みたいファイル
			is=new FileInputStream("C:\\Users\\onumaayano1199\\Pictures\\Sample1.xlsx");
			wb=WorkbookFactory.create(is);

			//どこのシート？
			Sheet sh =wb.getSheet("testsheet2");

			//どこの行？
			Row row =sh.getRow(0);

			//どこの列？
			Cell cell=row.getCell(0);

			//指定の値取ってきます
			value=cell.getStringCellValue();

			Row row1=sh.getRow(0);
			Cell cell1=row1.getCell(1);
			value1=cell1.getStringCellValue();


		}catch(Exception ex){
			ex.printStackTrace();
		}
		System.out.println(value);

		request.setAttribute("value", value);
		request.setAttribute("value1", value1);

		RequestDispatcher rd = request.getRequestDispatcher("/input.jsp");
		rd.forward(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}

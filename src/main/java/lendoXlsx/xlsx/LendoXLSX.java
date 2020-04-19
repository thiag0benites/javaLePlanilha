package lendoXlsx.xlsx;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import metodos.Excel;

/**
 * Hello world!
 *
 */
public class LendoXLSX 
{
	
	public static void main( String[] args )
    {
		Excel excel = new Excel();
		
		List<HashMap<String, String>> dados = new ArrayList<HashMap<String, String>>();
		dados = excel.leArquivoExcel("c:\\planilhas\\planilhaDaAula.xlsx");
		
		for (HashMap<String, String> dado : dados) {
			System.out.println(dado);
		}
    }
	
}

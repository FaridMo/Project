package MonPackage;

import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class ExcelReading {
		public static void main(String args){
			ArrayList<String> values = new ArrayList<String>();
			Connection connection;
			Statement statement; // statement = instruction il va nous permettre de pouvoir faire des requetes au niveau de notre base de données.
			
			try{
				
				InputStream file = new FileInputStream("users.xls");
				POIFSFileSystem pfs = new POIFSFileSystem(file);  // word; excel , powerpoint etc...
				HSSFWorkbook wb = new HSSFWorkbook(pfs);  // excel 
				HSSFSheet sheet = wb.getSheetAt(0);   // permet de spécifier le numero de la feuille sur lequel on travaille
				Iterator rows = sheet.rowIterator();   // permet de parcourir les lignes de la feuille de calcul
				
				while(rows.hasNext()){ // Tant qu'il y a des lignes !
					
					values.clear();
					
					HSSFRow row = (HSSFRow) rows.next();    // ligne suivante qui est la premiere ligne
					Iterator cells = row.cellIterator();    // parcourir les cellules dans chaque ligne
					
					while(cells.hasNext()){			// tant qu'il y a des cellules
						
						HSSFCell cell = (HSSFCell) cells.next(); // cellule suivante 
						
						if(HSSFCell.CELL_TYPE_NUMERIC == cell.getCellType()){
							values.add(Integer.toString((int) cell.getNumericCellValue())); //le type numérique est converti en entier puis en chaine de caractere car notre arraylist ne comporte que des chaines de caractères.
						}
						else if(HSSFCell.CELL_TYPE_STRING == cell.getCellType()){
								values.add(cell.getStringCellValue());
						}
						
					}
					
					try{
						Class.forName("com.mysql.jdbc.Driver").newInstance();
						connection = DriverManager.getConnection("jdbc.mysql://127.0.0.1/java","root","");
						statement = connection.createStatement();
						
						String sql = String
								.format("INSERT INTO users(first_name, last_name, username, password, age) VALUES('%s','%s','%s','%s','%S')",
										values.get(0), values.get(1),
										values.get(2), values.get(3),
										values.get(4) );
						
						int count = statement.executeUpdate(sql); //pour executer la requete qui est present au niveau de "sql" permet d'ajouter ou de modifier des infos
						statement.executeQuery(sql);
						
						/* 
							pour recupere des donnees d'une base de donnee on peut utiliser
							statement.executeQuery(sql);
					    */
						
						if(count > 0){
							System.out.println("Enregistrement effectuée !");
						}
						
						
					}
					
					
				}
				
			}
			
	}
}
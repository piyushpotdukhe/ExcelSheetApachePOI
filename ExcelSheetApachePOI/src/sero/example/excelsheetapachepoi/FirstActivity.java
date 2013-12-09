package sero.example.excelsheetapachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.os.Bundle;
import android.app.Activity;
import android.content.Context;
import android.content.res.AssetManager;
import android.util.Log;
import android.view.View;
import android.widget.Toast;

public class FirstActivity extends Activity {
	
	public final String FILE_NAME = "newxl.xls";

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.first_activity);
	} //e.o.onCreate

	
	
	
	
	//SEROTONIN: CREATE EXCELL FILE
	public void onClickCreateButton(View V) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("SHEET_ONE");
		
		Map<String, Object[]> map = new HashMap<String, Object[]>();
		Map<String, Object[]> data = new TreeMap<String, Object[]>(map); //sort map alphabetically
		data.put("1", new Object[] {"Emp No.", "Name", "Salary"});
		data.put("2", new Object[] {1d, "John", 1500000d});
		data.put("3", new Object[] {2d, "Sam", 800000d});
		data.put("4", new Object[] {3d, "Dean", 700000d});
		
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Log.v("SEROTONIN", "MAP: key=" + key);
		    Row row = sheet.createRow(rownum++);
		    Object [] objArr = data.get(key);
		    int cellnum = 0;
		    for (Object obj : objArr) {
		        Cell cell = row.createCell(cellnum++);
		        if(obj instanceof Date) 
		            cell.setCellValue((Date)obj);
		        else if(obj instanceof Boolean)
		            cell.setCellValue((Boolean)obj);
		        else if(obj instanceof String)
		            cell.setCellValue((String)obj);
		        else if(obj instanceof Double)
		            cell.setCellValue((Double)obj);
		    }
		}
		
		try {
			//File file_new = new File("/data/data/sero.example.excelsheetapachepoi/files/new.xls");
		    //FileOutputStream fos = new FileOutputStream(file_new);
		    FileOutputStream fos = openFileOutput(FILE_NAME, Context.MODE_WORLD_READABLE);
		    workbook.write(fos);
		    fos.close();
		    System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			Log.e("SEROTONIN", "EXCEPTION:FileNotFoundException: onClickCreateButton()-> create-write-save");
		    e.printStackTrace();
		} catch (IOException e) {
			Log.e("SEROTONIN", "EXCEPTION:IOException: onClickCreateButton()-> create-write-save");
		    e.printStackTrace();
		}
		
		
		
	} //e.o.onClickCreateButton
	
	
	
	
	
	
	
	
	//SEROTONIN: READ EXCEL FILE FROM ASSETS FOLDER.
	public void onClickReadFromAssetsButton(View V) {
		AssetManager assetManager = getResources().getAssets();
		InputStream inputStream = null;
		try {
		    inputStream = assetManager.open("SampleData.xls");
		    if (inputStream != null) {
		    	Log.d("SEROTONIN", "It worked!");
		    } else {
		    	Log.d("SEROTONIN", "FUCKED >i<");
		    }
		} catch (IOException e) {
			Log.e("SEROTONIN", "EXCEPTION:IOException: onClickReadFromAssetsButton()-> assetManager.open");
			e.printStackTrace();
		}

		Workbook workbook = null;
		try {
			workbook = new HSSFWorkbook(inputStream);
		} catch (FileNotFoundException e) {
			Log.e("SEROTONIN", "EXCEPTION:FileNotFoundException: onClickReadFromAssetsButton()-> HSSFWorkbook");
			e.printStackTrace();
		} catch (IOException e) {
			Log.e("SEROTONIN", "EXCEPTION:IOException: onClickReadFromAssetsButton()-> HSSFWorkbook");
			e.printStackTrace();
		}
		
		if (workbook != null) {
			Sheet sheet = workbook.getSheetAt(1);
			//READ
			try {
				Iterator<Row> rowIterator = sheet.iterator();
				while(rowIterator.hasNext()) {
					Row row = rowIterator.next();
					
					//For each row, iterate through each columns
				    Iterator<Cell> cellIterator = row.cellIterator();
				    while(cellIterator.hasNext()) {
				    	Cell cell = cellIterator.next();
				    	switch(cell.getCellType()) {
				    		case Cell.CELL_TYPE_BOOLEAN:
				    				System.out.print(cell.getBooleanCellValue() + "->");
				    				break;
				    		case Cell.CELL_TYPE_NUMERIC:
									System.out.print(cell.getNumericCellValue() + "->");
									break;
				    		case Cell.CELL_TYPE_STRING:
				    				System.out.print(cell.getStringCellValue() + "->");
				    				break;
				    	}
				    }
				    System.out.println("");
				}
				inputStream.close();
			} catch (FileNotFoundException e) {
				Log.e("SEROTONIN", "EXCEPTION:FileNotFoundException: onClickReadFromAssetsButton()-> READ");
				e.printStackTrace();
			} catch (IOException e) {
				Log.e("SEROTONIN", "EXCEPTION:IOException: onClickReadFromAssetsButton()-> READ");
				e.printStackTrace();
			}
		} else {
			Log.e("SEROTONIN", "onClickReadFromAssetsButton()-> workbook==null");
		}
	}//e.o.onClickReadFromAssetsButton

} //e.o.Activity
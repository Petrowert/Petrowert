package petrowert;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
	public static void main(String[] args) throws IOException {
		Integer i, j;
		String[] head_table = new String[15];
    	String[] inter = new String[15];
    	for(int a = 0; a < 15; a++)
    		inter[a] = "0";
    	Integer num_cur = 10; // Number current month
		Integer num_year = 2020; // Number current month
		Boolean y = false;
    	// Regular expression "\[[\d ]+:[\d \.]+];"
    	i = 2;
    	head_table[0] = "Number";
    	head_table[1] = "DZ";
    	
    	for(int k = 0; k < 13; k++) {
    		String zero;
    		if((num_cur + k - 1) % 12 < 10)
    			zero = "0";
    		else
    			zero = "";
    		if(y) {
    			head_table[k + 2] = String.valueOf(num_year) + zero + String.valueOf((num_cur + k - 1) % 12);
    			continue;
    		}
    		if((num_cur + k - 1) % 12 == 0) {
    			y = true;
    			head_table[k + 2] = String.valueOf(num_year) + String.valueOf(12);
    			continue;
    		}
    		head_table[k + 2] = String.valueOf(num_year - 1) + zero + String.valueOf((num_cur + k - 1) % 12);
    	}
    	FileWriter writer = new FileWriter("test_16k.csv");
    	for(String val: head_table) {
    		writer.append(val);
    		writer.append(';');
    	}
    	writer.append('\n');
    	try (InputStream inputStream = new FileInputStream(new File("data_16k.xlsx"))) { //FilePath from your device
	        Workbook workbook = StreamingReader.builder().rowCacheSize(200).bufferSize(4096).open(inputStream);
	        for (Sheet sheet : workbook) {
	        	i = 0;
	            for (Row row : sheet) {
	            	j = 0;
	                for (Cell cell : row) {
	                	// Filling table
	                    Pattern pattern = Pattern.compile("[\\d ]+:[-\\d \\.]+");
	                    Matcher matcher = pattern.matcher(getStringCellValue(cell));
	                    if(i > 0 && j < 2) {
	                    	inter[j] = getStringCellValue(cell);
	                    }
	                    while (matcher.find()) {
	                        String[] values = getStringCellValue(cell).substring(matcher.start(), matcher.end()).replace(" ", "").split(":");
	                        for(int b = 2; b < 15; b++) {
	                        	if(values[0].equals(head_table[b])) {
	                        		inter[b] = values[1];
	                        	}
	                        }
	                    }
	                	j++;
	                }
	                if(i > 0) {
		                for(int a = 0; a < 15; a++) {
		                	writer.append(inter[a]);
		                	writer.append(';');
		                }
		                writer.append('\n');
	                }
	                // Elements inter to zero.
					for(int a = 0; a < 15; a++)
	            		inter[a] = "0";
	                i++;
	            }
	        }
	        writer.flush();
	    	writer.close();
	        workbook.close();
	        inputStream.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}
	
	private static String getStringCellValue(Cell cell) {
		try {
	        switch (cell.getCellType()) {
	            case FORMULA:
	                try {
	                    return NumberToTextConverter.toText(cell.getNumericCellValue());
	                } catch (NumberFormatException e) {
	                    return cell.getStringCellValue();
	                }
	            case NUMERIC:
	                return NumberToTextConverter.toText(cell.getNumericCellValue());
	            case STRING:
	                String cellValue = cell.getStringCellValue().trim();
	                String pattern = "\\^\\$?-?([1-9][0-9]{0,2}(,\\d{3})*(\\.\\d{0,2})?|[1-9]\\d*(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^-?\\$?([1-9]\\d{0,2}(,\\d{3})*(\\.\\d{0,2})?|[1-9]\\d*(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^\\(\\$?([1-9]\\d{0,2}(,\\d{3})*(\\.\\d{0,2})?|[1-9]\\d*(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))\\)$";
	                if (((Pattern.compile(pattern)).matcher(cellValue)).find()) {
	                    return cellValue.replaceAll("[^\\d.]", "");
	                }
	                return cellValue.trim();
	            case BOOLEAN:
	                return String.valueOf(cell.getBooleanCellValue());
	            case ERROR:
	                return null;
	            default:
	                return cell.getStringCellValue();
	        }
	    } catch (Exception e) {
	        if (e.getLocalizedMessage() != null/* && ConfigReader.isDisplayWarnLog()*/)
	            return "";
	    }
	    return "";
	}
}

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
		Double summ = 0.;
    	Integer number_of_month = 12; // Original data count 13 month. From 092020 to 092020.
    	// Last payment in data 30.09.2020
    	Integer num_cur = 10; // Number current month
		Integer num_year = 2020; // Current year
		//Integer period = 3;
		Boolean y = false;
		String[] head_table = new String[number_of_month + 3];
    	String[] inter = new String[number_of_month + 3];
    	for(int a = 0; a < (number_of_month + 3); a++)
    		inter[a] = "0";
    	// Regular expression "\[[\d ]+:[\d \.]+];"
    	i = 2;
    	head_table[0] = "Number";
    	head_table[1] = "DZ";
    	head_table[2] = "MP"; // Average monthly payment. MP = (The amount of payments for the year)/12.
    	for(int k = 0; k < number_of_month; k++) {
    		String zero;
    		if((num_cur + k) % 12 < 10)
    			zero = "0";
    		else
    			zero = "";
    		if(y) {
    			head_table[k + 3] = String.valueOf(num_year) + zero + String.valueOf((num_cur + k) % 12);
    			continue;
    		}
    		if((num_cur + k) % 12 == 0) {
    			y = true;
    			head_table[k + 3] = String.valueOf(num_year - 1) + String.valueOf(12);
    			continue;
    		}
    		head_table[k + 3] = String.valueOf(num_year - 1) + zero + String.valueOf((num_cur + k) % 12);
    	}
    	/*FileWriter writer = new FileWriter("data_only1group.csv");
    	for(String val: head_table) {
    		writer.append(val);
    		writer.append(';');
    	}
    	writer.append("N group");
		writer.append(';');
    	writer.append('\n');
    	System.out.println();
    	try (InputStream inputStream = new FileInputStream(new File("data_cut_without_null.xlsx"))) { //FilePath from your device
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
	                        for(int b = 3; b < (number_of_month + 3); b++) {
	                        	if(values[0].equals(head_table[b])) {
	                        		inter[b] = values[1];
	                        		summ += Double.valueOf(values[1]);
	                        		//System.out.println("i = " + i + "; summ = " + "; summ = " + "; summ = " + summ);
	                        	}
	                        }
	                    }
	                	j++;
	                }
	                System.out.println();
	                // Check only 1 group of client
	                if(i > 0 && Double.valueOf(inter[1]) <= summ/12*1.06) {
	                	//System.out.println("summ = " + Double.valueOf(inter[1]) + "; summ = " + summ + "; summ/12 = " + summ/12 + "; summ/12*1.06 = " + summ/12*1.06);
	                	inter[2] = String.valueOf(Math.round(summ/12 * 100.0) / 100.0); // String.valueOf(summ);
		                for(int a = 0; a < (number_of_month + 3); a++) {
		                	writer.append(inter[a]);
		                	writer.append(';');
		                }
		                if(Double.valueOf(inter[1]) <= summ/12*1.06) {
		                	writer.append("1");
		                	writer.append(';');
		                }
		                writer.append('\n');
	                }
	                // Elements inter to zero.
					for(int a = 0; a < (number_of_month + 3); a++)
						inter[a] = "0";
	            	summ = 0.;
	                i++;
	            }
	        }
	        writer.flush();
	    	writer.close();
	        workbook.close();
	        inputStream.close();

    } catch (Exception e) {
        e.printStackTrace();
    }*/
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

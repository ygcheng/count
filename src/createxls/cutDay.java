package createxls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class cutDay {
	
	public static void writeExcel(OutputStream os) throws IOException, RowsExceededException, WriteException {
			WritableWorkbook wwb = Workbook.createWorkbook(os);
			WritableSheet ws = wwb.createSheet("Test Sheet 1", 0);
			for(int i = 0 ; i < 30 ; i ++ ){
				for(int j = 0 ; j < 6 ; j ++ ){
					Label label = new Label(j, i, "测试"+j+"-" + i);
					ws.addCell(label);
				}
			}
			
			WritableSheet ws2 = wwb.createSheet("Test Sheet 2", 1);
			for(int i = 0 ; i < 30 ; i ++ ){
				for(int j = 0 ; j < 6 ; j ++ ){
					Label label = new Label(j, i, "测试"+j+"-" + i);
					ws2.addCell(label);
				}
			}
			
			wwb.write();
			wwb.close();
		}
	
	public static void main(String[] args) throws Exception {
		File file = new File("E:/文档/Documents/舒浩哥");
		if(!file.exists()){
			file.mkdirs();
		}
		writeExcel(new FileOutputStream(file.getPath()+File.separator+ "hehe.xls"));
		
		/*
		Calendar calendar = Calendar.getInstance();
		
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		
		Date date1 = format.parse(format.format(calendar.getTime()));
		
		//31号之类的月末会出错
		calendar.set(Calendar.DAY_OF_MONTH,calendar.get(Calendar.DAY_OF_MONTH)+1);
		
		Date date2 = format.parse(format.format(calendar.getTime()));
		
		long jiange = date2.getTime() - date1.getTime();
		
		System.out.println("一天的间隔 : " + jiange);
		
		
		SimpleDateFormat f = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String d1 = "2015-11-02 07:00:00";
		String d2 = "2015-11-02 10:00:00";
		String d3 = "2015-11-02 12:00:00";
		String d4 = "2015-11-02 16:00:00";
		String d5 = "2015-11-02 23:00:00";
		String d6 = "2015-11-03 07:00:00";
		
		System.out.println("平-7点到10点-时间区间差值"+(f.parse(d2).getTime()-f.parse(d1).getTime()));
		System.out.println("峰-10点到12点-时间区间差值"+(f.parse(d3).getTime()-f.parse(d2).getTime()));
		System.out.println("平-12点到16点-时间区间差值"+(f.parse(d4).getTime()-f.parse(d3).getTime()));
		System.out.println("峰-16点到23点-时间区间差值"+(f.parse(d5).getTime()-f.parse(d4).getTime()));
		System.out.println("谷-23点到次日7点-时间区间差值"+(f.parse(d6).getTime()-f.parse(d5).getTime()));
		
		Date start = ExcelHandle.normalDateFormat.parse(format.format(new Date()) + " 07:00:00");
		
		System.out.println(start);*/
	}

}

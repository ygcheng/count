package createxls;

/**
 * 从excel文件C:/Users/Administrator/Desktop/6月二期.xls读取数据；
 * 生成新的excel文件C:/Users/Administrator/Desktop/6月二期2.xls
 * 修改原excel一个单元并输出为C:/Users/Administrator/Desktop/6月二期3.xls
 * @Description: TODO
 * @author 秭沐
 * @date 2015-10-29 下午3:23:10
 */

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ExcelHandle {

	public static SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm");
	
	public static SimpleDateFormat normalDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	/**
	 * 读取Excel
	 * @param filePath 文件路径 例如：E:/文档/Documents/舒浩哥/6月二期.xls
	 * @return List<Status> 把excel中的信息对象化并存入list中 ，，，（需求：真正有用时间段为启动到停止，，其余的状态可看着干扰信息，在这先行剔除）
	 */
	public static List<Status> readExcel(String filePath) {
		
		//原始数据使用list里存储
		List<Status> list = null;
		
		try {
			InputStream is = new FileInputStream(filePath);
			Workbook rwb = Workbook.getWorkbook(is);
			// 这里有两种方法获取sheet表:名字和下标（从0开始）
			// Sheet st = rwb.getSheet("original");
			Sheet st = rwb.getSheet(0);
			/**
			 * //获得第一行第一列单元的值 Cell c00 = st.getCell(0,0); //通用的获取cell值的方式,返回字符串
			 * String strc00 = c00.getContents(); //获得cell具体类型值的方式
			 * if(c00.getType() == CellType.LABEL) { LabelCell labelc00 =
			 * (LabelCell)c00; strc00 = labelc00.getString(); } //输出
			 * System.out.println(strc00);
			 */
			// Sheet的下标是从0开始
			// 获取第一张Sheet表
			Sheet rst = rwb.getSheet(0);
			// 获取Sheet表中所包含的总列数
			int rsColumns = rst.getColumns();
			// 获取Sheet表中所包含的总行数
			int rsRows = rst.getRows();
			// 获取指定单元格的对象引用
			
			list = new ArrayList<Status>();
			for (int i = 2; i < rsRows; i++) {
				
				Cell cell0 = rst.getCell(0, i);
				Cell cell1 = rst.getCell(1, i);
				Cell cell2 = rst.getCell(2, i);
				
				String name = cell0.getContents();
				if (name == null || "".equals(name)) {
					continue;
				}
				
				String operation = cell1.getContents().split(" ")[1];
				
				Date time = ExcelHandle.dateFormat.parse(cell2.getContents());
				Calendar calendar = Calendar.getInstance();
				calendar.setTime(time);
				calendar.set(Calendar.YEAR, calendar.get(Calendar.YEAR)+2000);
				
				Status status = new Status(name,operation,calendar.getTime());
				list.add(status);
				
				/*for (int j = 0; j < rsColumns; j++) {
					Cell cell = rst.getCell(j, i);
				}*/
			}
			// 关闭
			rwb.close();
			
			/*for(Status status : list){
				System.out.println(status.toString());
			}*/
			
			//System.out.println("***机柜运行记录 共"+list.size()+"条");
			//需求：真正有用时间段为启动到停止，，其余的状态可看着干扰信息，在这先行剔除
			for(int i = 0; i < list.size();){
				if(!(list.get(i).getOperation().equals("启动")||list.get(i).getOperation().equals("停止"))){
					list.remove(i);
				}else{
					i++;
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("读取excel出错");
		}
		
		return list;
	}

	/**
	 * 根据机柜的名称分组
	 * @param original 读取Excel的List<Status>
	 * @return List<List<Status>> 根据机柜的名称（status.getName）分组
	 */
	public static List<List<Status>> originalGrouping(List<Status> original) {

		List<List<Status>> group = new ArrayList<List<Status>>();
		
		for(Status status : original){
			String name = status.getName();
			if(0 == group.size()){
				
				List<Status> groupMember = new ArrayList<Status>();
				groupMember.add(status);
				group.add(groupMember);
				
			}else{
				
				boolean include = false;
				for(List<Status> groupMember:group){
					if(name.equals(groupMember.get(0).getName())){
						groupMember.add(status);
						include = true;
						break;
					}
				}
				if(!include){
					List<Status> groupMember = new ArrayList<Status>();
					groupMember.add(status);
					group.add(groupMember);
				}
				
			}
			
		}
		
		return group;
	}
	
	/**
	 * 找出同一组（也就是同一个机柜）正在运行的时间段TimeBucket。因为机器运行有停止的时候，所以出现多个TimeBucket
	 * @param group 根据机柜的名称分出的组List<List<Status>>
	 * @return List<List<TimeBucket>>
	 * @throws ParseException 时间在String和Date之间转化可能会抛出异常
	 */
	public static List<List<TimeBucket>> originalTB(List<List<Status>> group) throws ParseException {
		List<List<TimeBucket>> groupTB = new ArrayList<List<TimeBucket>>();
		
		
		for(List<Status> list : group){
			
			List<TimeBucket> listTB = ExcelHandle.getTimeBucket(list);
			
			groupTB.add(listTB);
		}
		
		return groupTB;
	}
	private static List<TimeBucket> getTimeBucket(List<Status> list) throws ParseException{
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		List<TimeBucket> listTB = new ArrayList<TimeBucket>();
		
		Status status0 = list.get(0);
		if("停止".equals(status0.getOperation())){
			listTB.add(new TimeBucket(null,status0,0));
		}else {
			Date time = format.parse(format.format(status0.getTime()));
			Calendar calendar = Calendar.getInstance();
			calendar.setTime(time);
			//月份加1有问题啊 。。。。。
			int month = calendar.get(Calendar.MONTH);
			if(month == 11){
				calendar.set(Calendar.YEAR, calendar.get(Calendar.YEAR));
				calendar.set(Calendar.MONTH, 1);
			}else{
				calendar.set(Calendar.MONTH, month + 1);
			}
			Status statusE = new Status(status0.getName(), "停止", calendar.getTime());
			listTB.add(new TimeBucket(status0,statusE,0));
		}
		
		for (int i = 1; i < list.size(); i++) {
			
			Status status = list.get(i);
			if("停止".equals(status.getOperation())){
				if(null != listTB.get(listTB.size()-1).getStart()){
					listTB.add(new TimeBucket(null,null,0));
					listTB.get(listTB.size()-1).setEnd(status);
				}
			}else {
				listTB.get(listTB.size()-1).setStart(status);
			}
			
		}
		
		if(null == listTB.get(listTB.size()-1).getStart()){
			Status statusE = listTB.get(listTB.size()-1).getEnd();
			Date time = format.parse(format.format(statusE.getTime()));
			listTB.get(listTB.size()-1).setStart(new Status(statusE.getName(), "启动", time));
		}
		
		return listTB;
	}
	
	/**
	 * 选出运行时间段（按峰平谷把时间段割小） TimeBucket分成更小的时间段且标注峰平谷
	 * 通过每天峰平谷的划分将运行时间段划分到更具体的峰平谷的小段
	 *   谷段：23:00-07:00
	 *   峰段：10:00-12:00 16:00-23:00
	 *   平段：07:00-10:00 12:00-16:00
	 * @param groupTB List<List<TimeBucket>>
	 * @return List<List<TimeBucket>>
	 * @throws ParseException 时间在String和Date之间转化可能会抛出异常
	 */
	public static List<List<TimeBucket>> splitTB(List<List<TimeBucket>> groupTB) throws ParseException {
		List<List<TimeBucket>> groupTB2 = new ArrayList<List<TimeBucket>>();
		for(List<TimeBucket> list1 : groupTB){
			groupTB2.add(new ArrayList<TimeBucket>());
			for(TimeBucket timeBucket : list1){
				groupTB2.get(groupTB2.size()-1).addAll(split(timeBucket));
			}
		}
		return groupTB2;
	}
	private static List<TimeBucket> split(TimeBucket timeBucket) throws ParseException{
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		List<TimeBucket> list = new ArrayList<TimeBucket>();
		/*平2-7点到10点-时间区间差值10800000 区间1
		峰1-10点到12点-时间区间差值7200000 区间2
		平2-12点到16点-时间区间差值14400000 区间3
		峰1-16点到23点-时间区间差值25200000 区间4
		谷3-23点到次日7点-时间区间差值28800000 区间5
		每个小时的时间区间差值3600000
		*/
		HashMap<Integer, Integer> trendMap = new HashMap<Integer, Integer>();
		trendMap.put(1, 2);
		trendMap.put(2, 1);
		trendMap.put(3, 2);
		trendMap.put(4, 1);
		trendMap.put(0, 3);
		
		//
		int nowQuJian = 5;
		Date startTime = ExcelHandle.normalDateFormat.parse(format.format(timeBucket.getStart().getTime()) + " 07:00:00");
		
		while (true) {
			
			if(startTime.getTime() <= timeBucket.getStart().getTime().getTime()){
				nowQuJian ++;
				
			}else if(startTime.getTime() <= timeBucket.getEnd().getTime().getTime()){
				TimeBucket temp = new TimeBucket();
				temp.setStart(new Status(timeBucket.getStart().getName(), timeBucket.getStart().getOperation(), timeBucket.getStart().getTime()));
				temp.setEnd(new Status(timeBucket.getStart().getName(), timeBucket.getStart().getOperation(), startTime));
				temp.setTrend(trendMap.get(nowQuJian%5));
				list.add(temp);
				
				timeBucket.setStart(new Status(timeBucket.getStart().getName(), timeBucket.getStart().getOperation(), startTime));
				
				nowQuJian ++;
			}else{
				TimeBucket temp = new TimeBucket();
				temp.setStart(new Status(timeBucket.getStart().getName(), timeBucket.getStart().getOperation(), timeBucket.getStart().getTime()));
				temp.setEnd(new Status(timeBucket.getEnd().getName(), timeBucket.getEnd().getOperation(), timeBucket.getEnd().getTime()));
				temp.setTrend(trendMap.get(nowQuJian%5));
				list.add(temp);
				break;
			}
			
			/*平2-7点到10点-时间区间差值10800000 区间1
			峰1-10点到12点-时间区间差值7200000 区间2
			平2-12点到16点-时间区间差值14400000 区间3
			峰1-16点到23点-时间区间差值25200000 区间4
			谷3-23点到次日7点-时间区间差值28800000 区间5
			每个小时的时间区间差值3600000
			*/
			switch (nowQuJian%5) {
			case 1:
				startTime = new Date(startTime.getTime()+10800000);
				break;
			case 2:
				startTime = new Date(startTime.getTime()+7200000);
				break;
			case 3:
				startTime = new Date(startTime.getTime()+14400000);
				break;
			case 4:
				startTime = new Date(startTime.getTime()+25200000);
				break;
			case 0:
				startTime = new Date(startTime.getTime()+28800000);
				break;
			default:
				break;
			}
			
		}
		
		return list;
	}
	
	
	// 测试
		public static void main(String args[]) throws IOException, ParseException {
				
			// 读EXCEL
			List<Status> original = ExcelHandle.readExcel("E:/文档/Documents/舒浩哥/6月二期.xls");
			
			if(original == null){
				System.out.println("读取excel失败");
				return;
			}
			
			//把原始数据分组
			List<List<Status>> group = ExcelHandle.originalGrouping(original);
			
			System.out.println("***原始数据分组 共" + group.size() + "组");
			/*for(int i = 0 ; i < group.size() ; i++ ){
				System.out.println("----第" +i+ "组");
				List<Status> list = group.get(i);
				for(Status status : list){
					System.out.println(status.toString());
				}
			}*/
				
			//选出运行时间段
			List<List<TimeBucket>> groupTB = ExcelHandle.originalTB(group);
			int i = 0;
			for(List<TimeBucket> list : groupTB){
				System.out.println("----第" +i+ "组");
				for(TimeBucket tb : list){
					System.out.println(tb.toString());
				}
				i++;
			}
			
			//选出运行时间段（按峰平谷把时间段割小）
			//通过每天峰平谷的划分将运行时间段划分到更具体的峰平谷的小段
					/*
					  谷段：23:00-07:00
					  峰段：10:00-12:00 16:00-23:00
					  平段：07:00-10:00 12:00-16:00
					 */
			/*List<List<TimeBucket>> groupTB2 = ExcelHandle.splitTB(groupTB);
			
			
			int i = 0;
			for(List<TimeBucket> list : groupTB2){
				float money1 = 0f;
				float money2 = 0f;
				float money3 = 0f;
			
				System.out.println("----第" +i+ "组");
				for(TimeBucket tb : list){
					System.out.println(tb.toString());
					//每个小时的时间区间差值3600000
					float hour = (tb.getEnd().getTime().getTime()-tb.getStart().getTime().getTime())/3600000.0f;
					if(tb.getTrend() == 1){
						money1 += hour*1.0192;
					}else if(tb.getTrend() == 2){
						money2 += hour*0.6162;
					}else if(tb.getTrend() == 3){
						money3 += hour*0.3062;
					}
				}
				System.out.println("----第" +i+ "组");
				System.out.println("峰：" + money1);
				System.out.println("平：" + money2);
				System.out.println("谷：" + money3);
				i++;
			}*/
			
		}


}
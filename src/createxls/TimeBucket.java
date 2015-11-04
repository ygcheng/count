/** 
* 2015-10-30 
* TimeBucket.java 
* author:秭沐 
*/

package createxls;

public class TimeBucket implements Comparable<TimeBucket>{
	
	private Status start;
	
	private Status end;
	
	/**
	 * 0表示还未分出峰平谷
	 * 1表示峰
	 * 2表示平
	 * 3表示谷
	 */
	private int trend;

	public Status getStart() {
		return start;
	}

	public void setStart(Status start) {
		this.start = start;
	}

	public Status getEnd() {
		return end;
	}

	public void setEnd(Status end) {
		this.end = end;
	}

	/**
	 * 
	 * @Description: TODO
	 * @param @return   
	 * @return int  0表示还未分出峰平谷  1表示峰  2表示平  3表示谷
	 * @throws
	 * @author 秭沐
	 * @date 2015-10-30 下午7:20:20
	 */
	public int getTrend() {
		return trend;
	}

	/**
	 * 0表示还未分出峰平谷  1表示峰  2表示平  3表示谷
	 * @Description: TODO
	 * @param @param trend  0表示还未分出峰平谷  1表示峰  2表示平  3表示谷
	 * @return void  
	 * @throws
	 * @author 秭沐
	 * @date 2015-10-30 下午7:18:40
	 */
	public void setTrend(int trend) {
		this.trend = trend;
	}

	/**
	 * 
	 * <p>Description: </p>
	 * @param start
	 * @param end
	 * @param trend  0表示还未分出峰平谷  1表示峰  2表示平  3表示谷
	 */
	public TimeBucket(Status start, Status end, int trend) {
		super();
		this.start = start;
		this.end = end;
		this.trend = trend;
	}
	public TimeBucket(){}

	@Override
	public String toString() {
		return "TimeBucket [start=" + start.getName()+ ":" + ExcelHandle.dateFormat.format(start.getTime())
				+ "("+ start.getOperation() +"), end=" + end.getName()+ ":" + ExcelHandle.dateFormat.format(end.getTime())
				+ "("+ end.getOperation() +"), trend=" + trend + ""+ getStr(trend) +"]"  + "-时间差：" + ((end.getTime().getTime()-start.getTime().getTime())/3600000f) ;
	}
	
	public String getStr(int trend){
		switch (trend) {
		case 0:
			return "(未处理)";
		case 1:
			return "(峰)";
		case 2:
			return "(平)";
        case 3:
        	return "(谷)";
		default:
			return "(未处理)";
		}
	}

	@Override
	public int compareTo(TimeBucket o) {
		if(this.start.getTime().after(o.getStart().getTime())){
			return -1;
		}else{
			return 1;
		}
		
	}

}

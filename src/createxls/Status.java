/** 
* 2015-10-30 
* Status.java 
* author:秭沐 
*/

package createxls;

import java.util.Date;

public class Status {
	

	private String name;
	
	private String operation;
	
	private Date time;
	
	

	public Status(String name, String operation, Date time) {
		super();
		this.name = name;
		this.operation = operation;
		this.time = time;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getOperation() {
		return operation;
	}

	public void setOperation(String operation) {
		this.operation = operation;
	}

	public Date getTime() {
		return time;
	}

	public void setTime(Date time) {
		this.time = time;
	}

	@Override
	public String toString() {
		return "Status [name=" + name + ", operation=" + operation + ", time="
				+ ExcelHandle.dateFormat.format(time) + "]";
	}
	

}

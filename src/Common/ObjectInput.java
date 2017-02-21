/*
 * Đối tượng tương ứng 1 cell trong excel
 * @Author: hieuht
 * @Date: 06/10/2016
 */
package Common;

public class ObjectInput {
	/*
	 * String id: Id trên website
	 * int type: loại (0-không xét, 1-luôn xét, 2-rỗng thì không xét)
	 * int acton: tương ứng với các action click, sendKey, ... 
	 * String name: tên cho mỗi đối tượng
	 * int index: số chỉ mục cho các đối tượng > 1 element trên web
	 */
	private String id;
	private int type;
	public int acton;
	public String name;
	public int index;
	public int timeSleep;
	public ObjectInput(String name, int type){
		this.type = type;
		this.name = name;
	}
	
	public ObjectInput(String name, String id){
		this.id = id;
		this.name = name;
	}
	
	public ObjectInput(String id, int action, int type, int index, int timeSleep) {
		ObjectInput(id, action, type, index, "", timeSleep);
	}
	
	private void ObjectInput(String id, int action, int type, int index, String name, int timeSleep) {
		this.id = id;
		this.type = type;
		this.acton = action;
		this.index = index;
		this.name = name;
		this.timeSleep = timeSleep;
	}

	public ObjectInput(String id, int action, int type, int index, String name, int timeSleep) {
		this.id = id;
		this.type = type;
		this.acton = action;
		this.index = index;
		this.name = name;
		this.timeSleep = timeSleep;
	}
	
	public String getId() {
		return id;
	}
	
	public void setId(String id) {
		this.id = id;
	}
	
	public int getType() {
		return type;
	}
	
	public void setType(int type) {
		this.type = type;
	}
	
	public int getActon() {
		return acton;
	}
	
	public void setActon(int acton) {
		this.acton = acton;
	}
	
	public String getName() {
		return name;
	}
	
	public void setName(String name) {
		this.name = name;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public int getTimeSleep() {
		return timeSleep;
	}

	public void setTimeSleep(int timeSleep) {
		this.timeSleep = timeSleep;
	}
	
}

package doqc;

public class DoQConverter {
	private int test_i=0;
	
	public DoQConverter(){
		this.test_i=-1;
	}
	
	public void calculate(){
		
	}
	
	public void beta(){
		this.test_i++;
	}
	
	public int info(){
		return this.test_i;
	}
	
	
	public static void main(String[] args){
		DoQConverter doqc=new DoQConverter();
		doqc.beta();
		doqc.beta();
		doqc.beta();
		
		System.out.println(doqc.info());
	}
}

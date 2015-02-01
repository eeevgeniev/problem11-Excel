package help;

import java.util.ArrayList;

public class Office implements Comparable<Office> {
	
	public String officeName;
	
	private ArrayList<Double> totalIncomes = new ArrayList<Double>();
	
	public void addIncome(double newIncome) {
		
		totalIncomes.add(newIncome);
		
	}
	
	public double returnTotalIncome() {
		
		double total = 0D;
		
		if (totalIncomes.size() != 0) {
			
			for (int i = 0; i < totalIncomes.size(); i++) {
				
				total += totalIncomes.get(i);
				
			}
			
		}
		
		return total;
		
	}

	@Override
	public int compareTo(Office otherOffice) {
		
		return this.officeName.compareTo(otherOffice.officeName);
		
	}

}

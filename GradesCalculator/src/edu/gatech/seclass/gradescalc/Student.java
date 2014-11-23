package edu.gatech.seclass.gradescalc;

import java.util.ArrayList;

public class Student {

	private String NAME;
	private String GTID;
	private String EMAIL;
	private int Attendance;
	private ArrayList<Number> IndividualGrades=new ArrayList<Number>();
	private ArrayList<Number> IndividualContribs=new ArrayList<Number>();
		
	public void setName(String NAME){
		this.NAME=NAME;
	}
	public void setGtid(String GTID){
		this.GTID=GTID;
	}
	public void setEmail(String EMAIL){
		this.EMAIL=EMAIL;
	}
	public void setAttendance(int Attendance){
		this.Attendance=Attendance;
	}
	public void setIndividualGrades(Number grade){
		IndividualGrades.add(grade);
	}
	public void setIndividualContribs(Number grade){
		IndividualContribs.add(grade);
	}
	public String getName(){
		return NAME;
	}
	
	public String getGtid(){
		return GTID;
	}
	
	public String getEmail(){
		return EMAIL;
	}
	
	public int getAttendance(){
		return Attendance;
	}
		
}

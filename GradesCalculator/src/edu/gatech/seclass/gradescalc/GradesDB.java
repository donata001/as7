package edu.gatech.seclass.gradescalc;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class GradesDB {
	 private XSSFSheet StudentsInfo,  Attendance, IndividualGrades, IndividualContribs, Teams,TeamGrades;
 	 private HashSet<Student> studentset=new HashSet<Student>();;
	 
	 public GradesDB(String FileDir){
 		 try
	        {
	            FileInputStream file = new FileInputStream(new File(FileDir));
	 
	            //Create Workbook instance holding reference to .xlsx file
	            XSSFWorkbook workbook = new XSSFWorkbook(file);
	 
	            //Get StudentInfo sheet from the workbook
	            StudentsInfo = workbook.getSheet("StudentsInfo");
	            addStudentInfo();
	            
	            //Get Attendance sheet from the workbook
	            Attendance=workbook.getSheet("Attendance");
	            addAttendance();
	            
	            //Get IndividualGrades sheet from the workbook
	            IndividualGrades=workbook.getSheet("IndividualGrades");
	            addIndividualGrades();
	            
	            //Get IndividualContribs sheet from the workbook
	            IndividualContribs=workbook.getSheet("IndividualContribs");
	            //addIndividualContribs();
	            
	            //Get Teams sheet from the workbook
	            Teams=workbook.getSheet("Teams");
	            
	            //Get TeamGrades sheet from the workbook
	            TeamGrades=workbook.getSheet("TeamGrades");
	            
	            file.close();
	        } 
	        catch (Exception e) 
	        {
	           e.printStackTrace();
	        }
	 }
	//add Student Info	 
	 private void addStudentInfo(){ 	
		 Iterator<Row> rowIterator = StudentsInfo.iterator();
	        while (rowIterator.hasNext()) 
	        {
	            Row row = rowIterator.next();
	            Student s=new Student();
	            s.setName(row.getCell(0).getStringCellValue());
	            if (row.getCell(1).getCellType()==0) 
	        	   s.setGtid(String.valueOf((long) row.getCell(1).getNumericCellValue()));
	            s.setEmail(row.getCell(2).getStringCellValue());
	            if (! s.getName().equals("NAME"))
	            		studentset.add(s);
	        }
	   }
	 //add Attendance for each student
	 private void addAttendance(){ 	
		 Iterator<Row> rowIterator2 = Attendance.iterator();
	        while (rowIterator2.hasNext()) 
	        {
	            Row row2 = rowIterator2.next();
	            for (Student s : studentset) {
	                if (s.getName().compareTo(row2.getCell(0).getStringCellValue()) == 0)
	                {
	                	 if (row2.getCell(1).getCellType()==0) 
	  		        	   s.setAttendance((int) row2.getCell(1).getNumericCellValue());
	                }
	           }        
	        }
	   }       
	 
	 //add Individual Grades for each student
	 private void addIndividualGrades(){ 	
		 Iterator<Row> rowIterator3 = IndividualGrades.iterator();
	        while (rowIterator3.hasNext()) 
	        {		 		            
	            Row row3 = rowIterator3.next();
	            for (Student s : studentset) {
	                if (s.getName().compareTo(row3.getCell(0).getStringCellValue()) == 0)
	                {
	                	if (row3.getCell(1).getCellType()==0) {
	  		        	    for (int i=1;i<IndividualGrades.getRow(0).getLastCellNum();i++)
	                		s.setIndividualGrades(row3.getCell(i).getNumericCellValue());
	 		            }
	                }
	           }        
	        }
	   }   
/*
	 //add Individual Contributes of each student to each project
	 private void addIndividualContribs() { 	
		 Iterator<Row> rowIterator4 = IndividualContribs.iterator();
	        while (rowIterator4.hasNext()) 
	        {		 		            
	        	Row row4 = rowIterator4.next();
	        	if (XSSFCell.CELL_TYPE_STRING==row4.getCell(0).getCellType() )
	        	{   
	        		for (Student s : studentset) {
	        	      if (s.getName().compareTo(row4.getCell(0).getStringCellValue()) == 0){
		                	if (row4.getCell(1).getCellType()==0) {
		  		        	    for (int i=1;i<IndividualGrades.getRow(0).getLastCellNum()-1;i++)
		                		s.setIndividualContribs(row4.getCell(i).getNumericCellValue());
		 		            }
		              }
	        	    }
	        	}
                else if (XSSFCell.CELL_TYPE_ERROR==row4.getCell(0).getCellType() )
                {
                    System.out.println("Unexpected Data Type");
                	break;
                }
	        }
	 }*/
	    
    // get Number of Students in the class
	public int getNumStudents(){
		return StudentsInfo.getLastRowNum();
	}
	
    // get Number of Individual Assignments 
	public int getNumAssignments(){
		return  IndividualGrades.getRow(0).getLastCellNum()-1; 
	}
	
    // get Number of Projects
	public int getNumProjects(){
		return IndividualContribs.getRow(0).getLastCellNum()-1; 
	}
	
    public HashSet<Student> getStudents(){
   	    return studentset;
    }
	
    // get one Student Class by Name
    public Student getStudentByName(String Name) throws IllegalArgumentException{
    	for (Student s : studentset) {
            if (s.getName().compareTo(Name) == 0)
              return s;
          }
            throw new IllegalArgumentException("no such student, check Name"); 
     }
    
    // get one Student Class by ID
    public Student getStudentByID(String GTID) throws IllegalArgumentException{
      	for (Student s : studentset) {
            if (s.getGtid().compareTo(GTID) == 0)
               return s;
        }
         throw new IllegalArgumentException("no such student, Check GTID"); 
        }
      
}

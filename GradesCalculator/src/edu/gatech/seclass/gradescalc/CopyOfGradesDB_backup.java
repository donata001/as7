package edu.gatech.seclass.gradescalc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyOfGradesDB_backup {
	 private XSSFSheet StudentsInfo,  Attendance, IndividualGrades, IndividualContribs, Teams,TeamGrades;
 	 private HashSet<Student> studentset=new HashSet<Student>();
 	 private XSSFWorkbook workbook;
 	 private String FileDir;
	 
	 // Initialize the Class
 	 public CopyOfGradesDB_backup(String FileDir){
		 this.FileDir=FileDir;
 		 try
	        {
 			    FileInputStream file = new FileInputStream(new File(FileDir));
	            //Create Workbook instance holding reference to .xlsx file
 			    workbook = new XSSFWorkbook(file);
	            //Get StudentInfo sheet from the workbook
	            StudentsInfo = workbook.getSheet("StudentsInfo");
	            readStudentInfo();
	            //Get Attendance sheet from the workbook
	            Attendance=workbook.getSheet("Attendance");
	            addAttendancetostudent();
	            //Get IndividualGrades sheet from the workbook
	            IndividualGrades=workbook.getSheet("IndividualGrades");
	            //Get IndividualContribs sheet from the workbook
	            IndividualContribs=workbook.getSheet("IndividualContribs");
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
	  	 
	 // Add New Assignment Name to the IndividualGrades sheet
	 public void addAssignment (String Assignment) throws IllegalArgumentException{
		    int lastcell=findcolumn(Assignment, IndividualGrades);
	        if (lastcell !=-1)    throw new IllegalArgumentException("Assignment already exits, check Assignment Name"); 
		    lastcell=IndividualGrades.getRow(0).getLastCellNum();
		    Cell cell = IndividualGrades.getRow(0).createCell(lastcell);
		    cell.setCellValue(Assignment);
		    // Update the Excel File
		     UpdateExcel("New Assignment Added");
	 }
	 
	 // Add New Project Name to the IndividualGrades sheet
	 public void addProject (String Project) throws IllegalArgumentException{
		    int lastcell=findcolumn(Project, IndividualContribs);
	        if (lastcell !=-1)    throw new IllegalArgumentException("Project already exits, check Assignment Name"); 
		    lastcell=IndividualContribs.getRow(0).getLastCellNum();
		    Cell cell = IndividualContribs.getRow(0).createCell(lastcell);
		    cell.setCellValue(Project);
		    
		    int lastcell2=findcolumn(Project, TeamGrades);
	        if (lastcell2 !=-1)    throw new IllegalArgumentException("Project already exits, check Assignment Name"); 
		    lastcell2=TeamGrades.getRow(0).getLastCellNum();
		    Cell cell2 = TeamGrades.getRow(0).createCell(lastcell2);
		    cell2.setCellValue(Project);
			    
		    // Update the Excel File
		     UpdateExcel("New Project Added");		 
	 }
	 
	 // Add New Teams 
	 public void addTeam (String TeamName) throws IllegalArgumentException{
		  int rownum=Teams.getPhysicalNumberOfRows()-1; 
    	  for (String newteam: TeamNames){
		  Row newrow=Teams.createRow(rownum);
	      newrow.createCell(0).setCellValue(newteam);
	      Row newrow2=TeamGrades.createRow(rownum);
	      newrow2.createCell(0).setCellValue(newteam);
	      rownum++;
    	  }
	      UpdateExcel("New Team Added"); 	    
   	    	 }
	 
	 // Add New Grades to corresponding Assignment in the IndividualGrades sheet
	 public void addGradesForAssignment(String assignmentName, HashMap<Student, Integer> grades) throws IllegalArgumentException{
		 // find the corresponding Assignment	   
		 int cellnum=findcolumn(assignmentName,IndividualGrades);
	     if (cellnum==-1)    throw new IllegalArgumentException("no such Assignment, check Assignment Name"); 

	     Set<Student> keyset = grades.keySet();
	     for (Student key : keyset) {
	          int rownum=findrow(key.getName(),IndividualGrades);
		      if (rownum==-1)    throw new IllegalArgumentException("no such Student, check Student Name"); 

	      // write the grades to the corresponding Assignment and Student
	      Row row = IndividualGrades.getRow(rownum);
          Cell assignmentgrades = row.createCell(cellnum);
          assignmentgrades.setCellValue(grades.get(key));
	      }
	      // Update the Excel File
	      UpdateExcel("New Individual Assignment Grades Added" );
		 
	 }
	 // Add New individual contributions to corresponding Project in the IndividuaContribs sheet
	 public void addIndividualContributions(String projectName, HashMap<Student, Integer> contributions) throws IllegalArgumentException{
		 int cellnum=findcolumn(projectName,IndividualContribs);
	     if (cellnum==-1)    throw new IllegalArgumentException("no such Project, check Project Name"); 
		 Set<Student> keyset = contributions.keySet();
	     for (Student key : keyset) {
	          int rownum=findrow(key.getName(),IndividualContribs);
		 if (rownum==-1)    throw new IllegalArgumentException("no such Student, check Student Name"); 
	      // write the grades to the corresponding Assignment and Student
	      Row row = IndividualContribs.getRow(rownum);
	         Cell projectcontribs = row.createCell(cellnum);
	         projectcontribs.setCellValue(contributions.get(key));
	      }
	      // Update the Excel File
	      UpdateExcel("New Individual Contributions to Project Added" );
		 }
	 
	 // Add Attendance to Attendance sheet
	 public void addAttendance(HashMap<Student, Integer> attendance) throws IllegalArgumentException{
		 Set<Student> keyset = attendance.keySet();
	     for (Student key : keyset) {
	          int rownum=findrow(key.getName(),Attendance);
		 if (rownum==-1)    throw new IllegalArgumentException("no such Student, check Student Name"); 
	      // write the attendance to the corresponding Student
	      Row row = Attendance.getRow(rownum);
	         Cell projectcontribs = row.createCell(1);
	         projectcontribs.setCellValue(attendance.get(key));
	      }
	      // Update the Excel File
	      UpdateExcel("Attendance Added");
		 }
	 
	 // Add Student to Team
	 public void addStudentToTeam(HashMap<Student, String> team) throws IllegalArgumentException{
		 Set<Student> keyset = team.keySet();
	     for (Student key : keyset) {
	          int rownum=findrow(team.get(key),Teams);
	          if (rownum==-1){
	        	  throw new IllegalArgumentException("no such Team, check Team Name"); 
		        }
	           Row row = Teams.getRow(rownum);
	           int cellnum=row.getPhysicalNumberOfCells();
	           Cell projectcontribs = row.createCell(cellnum);
	           projectcontribs.setCellValue(key.getName());
	      }
	      // Update the Excel File
	      UpdateExcel("Student Added to Team");
	      }
	 
	 // Add team grades to corresponding Project in the TeamGrades sheet
	 public void addTeamGradesForProject(String projectName, HashMap<String, Integer> teamgrades) throws IllegalArgumentException{
	     // find the corresponding Assignment	   
		 int cellnum=findcolumn(projectName,TeamGrades);
	     if (cellnum==-1)    throw new IllegalArgumentException("no such Project, check Project Name"); 

		 Set<String> keyset = teamgrades.keySet();
	     for (String key : keyset) {
	          int rownum=findrow(key,TeamGrades);
		      if (rownum==-1)    throw new IllegalArgumentException("no such Team Name, check Team Name"); 
	          // write the grades to the corresponding Assignment and Student
	          Row row = TeamGrades.getRow(rownum);
	          Cell projectgrades = row.createCell(cellnum);
	          projectgrades.setCellValue(teamgrades.get(key));
	      }
	      // Update the Excel File
	      UpdateExcel("New Teamgrades to Project Added");
		 }
	 
	 // Add New students
	 public void addStudent(HashSet<Student> newstudents) throws IllegalArgumentException{
	     for (Student newstudent : newstudents) {
			 // add Student into StudentsInfo sheet
	    	 int rownum=findrow(newstudent.getName(),StudentsInfo);
			    if (rownum!=-1)    throw new IllegalArgumentException("this student already exits in StudentsInfo sheet"); 
	         rownum=StudentsInfo.getPhysicalNumberOfRows();
	         Row row1 = StudentsInfo.createRow(rownum);
	         Row row2 = Attendance.createRow(rownum);
	         Row row3 = IndividualGrades.createRow(rownum);
	         Row row4 = IndividualContribs.createRow(rownum);

	         Cell cellname1 = row1.createCell(0);
	         Cell cellname2 = row2.createCell(0);
	         Cell cellname3 = row3.createCell(0);
	         Cell cellname4 = row4.createCell(0);

	         cellname1.setCellValue(newstudent.getName());
	         cellname2.setCellValue(newstudent.getName());
	         cellname3.setCellValue(newstudent.getName());
	         cellname4.setCellValue(newstudent.getName());

	         Cell cellGtid= row1.createCell(1);
	         cellGtid.setCellValue( Long.parseLong(newstudent.getGtid()));
         
	      }
	      // Update the Excel File
	      UpdateExcel("New Student Added");
		 }

	 
	 //get the average grades of all assignments for specific student
	 public int getAverageAssignmentsGrade(Student student) throws IllegalArgumentException{
		  CheckSheet(student.getName(),IndividualGrades, "Missing IndividualGrades for this student, Please Check IndividualGrades");
		  // find the corresponding student
          int rownum=findrow(student.getName(),IndividualGrades );
	      double sum=0;
	      for (int i=1; i<IndividualGrades.getRow(rownum).getLastCellNum();i++ )
	    	  sum+=IndividualGrades.getRow(rownum).getCell(i).getNumericCellValue();
	      
		 return (int) Math.ceil(sum / (IndividualGrades.getRow(rownum).getLastCellNum()-1));
	 }
	 
	 //get the average grades of all Projects for specific student
	 public int getAverageProjectsGrade(Student student) throws IllegalArgumentException{
		 CheckSheet(student.getName(),IndividualContribs, "Missing IndividualGrades for this student, Please Check IndividualGrades");
		 // find the corresponding student
         int rownum=findrow(student.getName(),IndividualContribs);
         // find team by name:
		 String team=getTeambyName(student.getName());
		 CheckSheet(team,TeamGrades, "Missing TeamGrades for this team, Please Check TeamGrades");

		 int projectgrades=0;
		 for(int i=1; i<IndividualContribs.getRow(rownum).getPhysicalNumberOfCells(); i++){
			 double teamgrades= TeamGrades.getRow(findrow(team,TeamGrades)).getCell(i).getNumericCellValue();
			 double contribs=IndividualContribs.getRow(rownum).getCell(i).getNumericCellValue();
			 projectgrades+=contribs*teamgrades;
		 }
	      return (int) Math.ceil(projectgrades/100/(IndividualContribs.getRow(rownum).getPhysicalNumberOfCells()-1));
	 }
	    
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
    
//*******************************************************
//******Following is used privately by the class*********
//*******************************************************
	//read Student Info	from datasheet
	 private void readStudentInfo(){ 	
		 Iterator<Row> rowIterator = StudentsInfo.iterator();
	        while (rowIterator.hasNext()) 
	        {
	            Row row = rowIterator.next();
	            Student s=new Student(null,null,null);
	            s.setName(row.getCell(0).getStringCellValue());
	            if (row.getCell(1).getCellType()==0) 
	        	   s.setGtid(String.valueOf((long) row.getCell(1).getNumericCellValue()));
	            //s.setEmail(row.getCell(2).getStringCellValue());
	            if (! s.getName().equals("NAME"))
	            		studentset.add(s);
	        }
	   }

	 
	 
	 
	 // this method is used to find the corresponding row number by student name in given sheet
	 private int findrow (String NAME, XSSFSheet sheet) throws IllegalArgumentException{
		 int rownum =-1;
	     for (int i=0;i<sheet.getPhysicalNumberOfRows();i++){
	    	 //System.out.println(sheet.getRow(i).getCell(0).getCellType());
	    	 //System.out.println(sheet.getRow(i).getCell(0).getStringCellValue());
	    	 switch(sheet.getRow(i).getCell(0).getCellType()) {
             case Cell.CELL_TYPE_ERROR:
	              break;      
             case Cell.CELL_TYPE_STRING: 
	             if (NAME.equals(sheet.getRow(i).getCell(0).getStringCellValue()))
		           {rownum=i;
		           return rownum;}
             case Cell.CELL_TYPE_BLANK:
            	   break;    
	             
		     }
	     }
	     return rownum;
	 }
	 
	 private int findcolumn(String toFound, XSSFSheet sheet) throws IllegalArgumentException{
		 int cellnum=-1; 
		 for (int i=0;i<sheet.getRow(0).getPhysicalNumberOfCells();i++){
		      if (toFound.equals(sheet.getRow(0).getCell(i).getStringCellValue()))
		       {cellnum=i;
		       break;}
		 }
		 return cellnum;
	 }
	 
	 public String getTeambyName(String name) throws IllegalArgumentException{
		 String team =null;
		 Iterator<Row> rowIterator = Teams.iterator();
	     while (rowIterator.hasNext()){
	         Row row = rowIterator.next();
	    	 Iterator<Cell> cellIterator = row.cellIterator();
	           while(cellIterator.hasNext()) {
 	                Cell cell = cellIterator.next();
	    	        if (name.equals(cell.getStringCellValue())){
	    	        	team=row.getCell(0).getStringCellValue();
	    	        	return team;
	    	        }
	          }
	     }
	     throw new IllegalArgumentException("Cannot find the team for this student");  
	 }
	 
	 private void UpdateExcel(String message){
     try {
		    FileOutputStream outFile =new FileOutputStream(new File(FileDir));
		    workbook.write(outFile);
		    outFile.close();
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
          System.out.println(message+ " Successfully!");
	 }
	 
	 private void CheckSheet(String item, XSSFSheet sheet, String Message) throws IllegalArgumentException{
		 int rownum=findrow(item,sheet);
		    int columnoftitle =sheet.getRow(0).getPhysicalNumberOfCells();
	        for (int i=0; i<columnoftitle;i++){
	            Cell cell = sheet.getRow(rownum).getCell(i);
	            if (cell==null)
                	throw new IllegalArgumentException(Message); 
	         }
      }
	 
	 //add Attendance for each student class
	 private void addAttendancetostudent(){ 	
		 Iterator<Row> rowIterator2 = Attendance.iterator();
	        while (rowIterator2.hasNext()) 
	        {
	            Row row2 = rowIterator2.next();
	            for (Student s : studentset) {
	                if (s.getName().compareTo(row2.getCell(0).getStringCellValue()) == 0)
	                {	Cell cell=row2.getCell(1);
	                    if (cell==null)
	                	      break;  //throw new IllegalArgumentException("Excepted Null row"); 
	                     else
		  		        	  s.setAttendance((int) cell.getNumericCellValue());
	                }
	           }        
	        }
	   }   
}

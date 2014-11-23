package edu.gatech.seclass.gradescalc;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class CopyOfGradesDBTest2 {

    GradesDB db = null;
    static final String GRADES_DB_GOLDEN = "DB" + File.separator
            + "GradesDatabase6300-goldenversion.xlsx";
    static final String GRADES_DB = "DB" + File.separator
            + "GradesDatabase6300.xlsx";

    @Before
    public void setUp() throws Exception {
        FileSystem fs = FileSystems.getDefault();
        Path dbfilegolden = fs.getPath(GRADES_DB_GOLDEN);
        Path dbfile = fs.getPath(GRADES_DB);
        Files.copy(dbfilegolden, dbfile, REPLACE_EXISTING);
        db = new GradesDB(GRADES_DB);
    }

    @After
    public void tearDown() throws Exception {
        db = null;
    }


    @Test
    public void testAddStudentInfo() {
        Student student1 = new Student("James Lebran", "911234502", db);
        Student student2 = new Student("Chris Paul", "911234508", db);
        Student student3 = new Student("James Harden", "911234510", db);
        
        HashSet<Student> newstudents = new HashSet<Student>();
        newstudents.add(student1);
        newstudents.add(student2);
        newstudents.add(student3);
        db.addStudent(newstudents);
        
        student1.setEmail("jl@gatech.edu");
        student1.setGtid("1000001");
        db.addStudentInfo(student1);
        db = new GradesDB(GRADES_DB);
        Student student = db.getStudentByName("James Lebran");
        assertTrue(student.getGtid().compareTo("1000001") == 0);
        System.out.println(student.getEmail());
        assertTrue(student.getEmail().compareTo("jl@gatech.edu") == 0);

        
    }

}

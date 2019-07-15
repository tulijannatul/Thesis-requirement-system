package thesis_requirements;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Thesis_requirements 
{

    public static void main(String[] args) throws Exception 
    {
        
        File f = new File("C:\\Users\\User\\Desktop\\250\\Requirements for thesis (Responses).xlsx");
        
        FileInputStream fis = new FileInputStream(f);
        
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        
        XSSFSheet sh1 = wb.getSheetAt(0);
        
        int rownum = sh1.getLastRowNum();
        
        int flag1, flag2, j, m, n;
        m = 2;//for using Flag
        
        //System.out.println(rownum);
        
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\User\\Desktop\\250\\Test.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet worksheet = workbook.createSheet("Newsheet");
        
        XSSFRow row1 = worksheet.createRow(0);
        
        XSSFCell cellA1 = row1.createCell(0);
        cellA1.setCellValue("Name of member 1");
        XSSFCell cellB1 = row1.createCell(1);
        cellB1.setCellValue("Reg. No. of member 1");
        XSSFCell cellC1 = row1.createCell(2);
        cellC1.setCellValue("Name of member 2");
        XSSFCell cellD1 = row1.createCell(3);
        cellD1.setCellValue("Reg. No. of member 2");
        XSSFCell cellE1 = row1.createCell(4);
        cellE1.setCellValue("Structured Programming Language of member 1");
        XSSFCell cellF1 = row1.createCell(5);
        cellF1.setCellValue("Structured Programming Language Lab of member 1");
        XSSFCell cellG1 = row1.createCell(6);
        cellG1.setCellValue("Data Structure of member 1");
        XSSFCell cellH1 = row1.createCell(7);
        cellH1.setCellValue("Data Structure Lab of member 1");
        XSSFCell cellI1 = row1.createCell(8);
        cellI1.setCellValue("Discrete Mathematics of member 1");
        XSSFCell cellJ1 = row1.createCell(9);
        cellJ1.setCellValue("Discrete Mathematics Lab of member 1");
        XSSFCell cellK1 = row1.createCell(10);
        cellK1.setCellValue("Object Oriented Programming Language of member 1");
        XSSFCell cellL1 = row1.createCell(11);
        cellL1.setCellValue("Object Oriented Programming Language Lab of member 1");
        XSSFCell cellM1 = row1.createCell(12);
        cellM1.setCellValue("Algorithm Design and Analysis of member 1");
        XSSFCell cellN1 = row1.createCell(13);
        cellN1.setCellValue("Algorithm Design and Analysis Lab of member 1");
        XSSFCell cellO1 = row1.createCell(14);
        cellO1.setCellValue("Database System of member 1");
        XSSFCell cellP1 = row1.createCell(15);
        cellP1.setCellValue("Database System Lab of member 1");
        XSSFCell cellQ1 = row1.createCell(16);
        cellQ1.setCellValue("Operating System and System Programming of member 1");
        XSSFCell cellR1 = row1.createCell(17);
        cellR1.setCellValue("Operating System and System Programming Lab of member 1");
        XSSFCell cellS1 = row1.createCell(18);
        cellS1.setCellValue("Numerical Analysis of member 1");
        XSSFCell cellT1 = row1.createCell(19);
        cellT1.setCellValue("Numerical Analysis Lab of member 1");
        XSSFCell cellU1 = row1.createCell(20);
        cellU1.setCellValue("Theory of Computation of mamber 1");
        XSSFCell cellV1 = row1.createCell(21);
        cellV1.setCellValue("Computer Networking of mamber 1");
        XSSFCell cellW1 = row1.createCell(22);
        cellW1.setCellValue("Computer Networking Lab of mamber 1");
        XSSFCell cellX1 = row1.createCell(23);
        cellX1.setCellValue("Computer Graphics and Image Processing of mamber 1");
        XSSFCell cellY1 = row1.createCell(24);
        cellY1.setCellValue("Computer Graphics and Image Processing Lab of mamber 1");
        XSSFCell cellZ1 = row1.createCell(25);
        cellZ1.setCellValue("Technical Writing and Presentation of member 1");
        XSSFCell cellAA1 = row1.createCell(26);
        cellAA1.setCellValue("Project 150 of member 1");
        XSSFCell cellAB1 = row1.createCell(27);
        cellAB1.setCellValue("Project 250 of member 1");
        XSSFCell cellAC1 = row1.createCell(28);
        cellAC1.setCellValue("Project 350 of member 1");
        
        XSSFCell cellAE1 = row1.createCell(30);
        cellAE1.setCellValue("Structured Programming Language of member 2");
        XSSFCell cellAF1 = row1.createCell(31);
        cellAF1.setCellValue("Structured Programming Language Lab of member 2");
        XSSFCell cellAG1 = row1.createCell(32);
        cellAG1.setCellValue("Data Structure of member 2");
        XSSFCell cellAH1 = row1.createCell(33);
        cellAH1.setCellValue("Data Structure Lab of member 2");
        XSSFCell cellAI1 = row1.createCell(34);
        cellAI1.setCellValue("Discrete Mathematics of member 2");
        XSSFCell cellAJ1 = row1.createCell(35);
        cellAJ1.setCellValue("Discrete Mathematics Lab of member 2");
        XSSFCell cellAK1 = row1.createCell(36);
        cellAK1.setCellValue("Object Oriented Programming Language of member 2");
        XSSFCell cellAL1 = row1.createCell(37);
        cellAL1.setCellValue("Object Oriented Programming Language Lab of member 2");
        XSSFCell cellAM1 = row1.createCell(38);
        cellAM1.setCellValue("Algorithm Design and Analysis of member 2");
        XSSFCell cellAN1 = row1.createCell(39);
        cellAN1.setCellValue("Algorithm Design and Analysis Lab of member 2");
        XSSFCell cellAO1 = row1.createCell(40);
        cellAO1.setCellValue("Database System of member 2");
        XSSFCell cellAP1 = row1.createCell(41);
        cellAP1.setCellValue("Database System Lab of member 2");
        XSSFCell cellAQ1 = row1.createCell(42);
        cellAQ1.setCellValue("Operating System and System Programming of member 2");
        XSSFCell cellAR1 = row1.createCell(43);
        cellAR1.setCellValue("Operating System and System Programming Lab of member 2");
        XSSFCell cellAS1 = row1.createCell(44);
        cellAS1.setCellValue("Numerical Analysis of member 2");
        XSSFCell cellAT1 = row1.createCell(45);
        cellAT1.setCellValue("Numerical Analysis Lab of member 2");
        XSSFCell cellAU1 = row1.createCell(46);
        cellAU1.setCellValue("Theory of Computation of mamber 2");
        XSSFCell cellAV1 = row1.createCell(47);
        cellAV1.setCellValue("Computer Networking of mamber 2");
        XSSFCell cellAW1 = row1.createCell(48);
        cellAW1.setCellValue("Computer Networking Lab of mamber 2");
        XSSFCell cellAX1 = row1.createCell(49);
        cellAX1.setCellValue("Computer Graphics and Image Processing of mamber 2");
        XSSFCell cellAY1 = row1.createCell(50);
        cellAY1.setCellValue("Computer Graphics and Image Processing Lab of mamber 2");
        XSSFCell cellAZ1 = row1.createCell(51);
        cellAZ1.setCellValue("Technical Writing and Presentation of member 2");
        XSSFCell cellBA1 = row1.createCell(52);
        cellBA1.setCellValue("Project 150 of member 2");
        XSSFCell cellBB1 = row1.createCell(53);
        cellBB1.setCellValue("Project 250 of member 2");
        XSSFCell cellBC1 = row1.createCell(54);
        cellBC1.setCellValue("Project 350 of member 2");
        
        
        
         
        for(int i = 1; i <= rownum; i++)
        {
            //Member 1 data input
            
            String data1 = sh1.getRow(i).getCell(1).getStringCellValue(); //Name
            int data2 = (int) sh1.getRow(i).getCell(2).getNumericCellValue(); //Reg
            String data3 = sh1.getRow(i).getCell(3).getStringCellValue();  //Mail
            double data4 = sh1.getRow(i).getCell(4).getNumericCellValue(); //Mobile
            double data5 = sh1.getRow(i).getCell(5).getNumericCellValue(); //Total_credit
            double data6 = sh1.getRow(i).getCell(6).getNumericCellValue(); //CGPA
            
            String data7 = sh1.getRow(i).getCell(7).getStringCellValue(); //Course
            String data8 = sh1.getRow(i).getCell(8).getStringCellValue(); //Course
            String data9 = sh1.getRow(i).getCell(9).getStringCellValue(); //Course
            String data10 = sh1.getRow(i).getCell(10).getStringCellValue(); //Course
            String data11 = sh1.getRow(i).getCell(11).getStringCellValue(); //Course
            String data12 = sh1.getRow(i).getCell(12).getStringCellValue(); //Course
            String data13 = sh1.getRow(i).getCell(13).getStringCellValue(); //Course
            String data14 = sh1.getRow(i).getCell(14).getStringCellValue(); //Course
            String data15 = sh1.getRow(i).getCell(15).getStringCellValue(); //Course
            String data16 = sh1.getRow(i).getCell(16).getStringCellValue(); //Course
            String data17 = sh1.getRow(i).getCell(17).getStringCellValue(); //Course
            String data18 = sh1.getRow(i).getCell(18).getStringCellValue(); //Course
            String data19 = sh1.getRow(i).getCell(19).getStringCellValue(); //Course
            String data20 = sh1.getRow(i).getCell(20).getStringCellValue(); //Course
            String data21 = sh1.getRow(i).getCell(21).getStringCellValue(); //Course
            String data22 = sh1.getRow(i).getCell(22).getStringCellValue(); //Course
            String data23 = sh1.getRow(i).getCell(23).getStringCellValue(); //Course
            String data24 = sh1.getRow(i).getCell(24).getStringCellValue(); //Course
            String data25 = sh1.getRow(i).getCell(25).getStringCellValue(); //Course
            String data26 = sh1.getRow(i).getCell(26).getStringCellValue(); //Course
            String data27 = sh1.getRow(i).getCell(27).getStringCellValue(); //Course
            String data28 = sh1.getRow(i).getCell(28).getStringCellValue(); //Course
            String data29 = sh1.getRow(i).getCell(29).getStringCellValue(); //Course
            String data30 = sh1.getRow(i).getCell(30).getStringCellValue(); //Course
            String data31 = sh1.getRow(i).getCell(31).getStringCellValue(); //Course
            
            
            //Member 2 data input
            
            String data32 = sh1.getRow(i).getCell(32).getStringCellValue(); //Name
            int data33 = (int) sh1.getRow(i).getCell(33).getNumericCellValue(); //Reg
            String data34 = sh1.getRow(i).getCell(34).getStringCellValue();  //Mail
            double data35 = sh1.getRow(i).getCell(35).getNumericCellValue(); //Mobile
            double data36 = sh1.getRow(i).getCell(36).getNumericCellValue(); //Total_credit
            double data37 = sh1.getRow(i).getCell(37).getNumericCellValue(); //CGPA
            
            String data38 = sh1.getRow(i).getCell(38).getStringCellValue(); //Course
            String data39 = sh1.getRow(i).getCell(39).getStringCellValue(); //Course
            String data40 = sh1.getRow(i).getCell(40).getStringCellValue(); //Course
            String data41 = sh1.getRow(i).getCell(41).getStringCellValue(); //Course
            String data42 = sh1.getRow(i).getCell(42).getStringCellValue(); //Course
            String data43 = sh1.getRow(i).getCell(43).getStringCellValue(); //Course
            String data44 = sh1.getRow(i).getCell(44).getStringCellValue(); //Course
            String data45 = sh1.getRow(i).getCell(45).getStringCellValue(); //Course
            String data46 = sh1.getRow(i).getCell(46).getStringCellValue(); //Course
            String data47 = sh1.getRow(i).getCell(47).getStringCellValue(); //Course
            String data48 = sh1.getRow(i).getCell(48).getStringCellValue(); //Course
            String data49 = sh1.getRow(i).getCell(49).getStringCellValue(); //Course
            String data50 = sh1.getRow(i).getCell(50).getStringCellValue(); //Course
            String data51 = sh1.getRow(i).getCell(51).getStringCellValue(); //Course
            String data52 = sh1.getRow(i).getCell(52).getStringCellValue(); //Course
            String data53 = sh1.getRow(i).getCell(53).getStringCellValue(); //Course
            String data54 = sh1.getRow(i).getCell(54).getStringCellValue(); //Course
            String data55 = sh1.getRow(i).getCell(55).getStringCellValue(); //Course
            String data56 = sh1.getRow(i).getCell(56).getStringCellValue(); //Course
            String data57 = sh1.getRow(i).getCell(57).getStringCellValue(); //Course
            String data58 = sh1.getRow(i).getCell(58).getStringCellValue(); //Course
            String data59 = sh1.getRow(i).getCell(59).getStringCellValue(); //Course
            String data60 = sh1.getRow(i).getCell(60).getStringCellValue(); //Course
            String data61 = sh1.getRow(i).getCell(61).getStringCellValue(); //Course
            String data62 = sh1.getRow(i).getCell(62).getStringCellValue(); //Course
            
           
            flag1 = 0;
            flag2 = 0;
            
            if ((data6 >= 3.00) && 
                    (!data7.equals("F")) && (!data8.equals("F")) && (!data9.equals("F")) && (!data10.equals("F")) && (!data11.equals("F")) && (!data12.equals("F")) && (!data13.equals("F")) && (!data14.equals("F")) && (!data15.equals("F")) && (!data16.equals("F")) && (!data17.equals("F")) && (!data18.equals("F")) &&(!data19.equals("F")) && (!data20.equals("F")) && (!data21.equals("F")) && (!data22.equals("F")) && (!data23.equals("F")) && (!data24.equals("F")) && (!data25.equals("F")) && (!data26.equals("F")) && (!data27.equals("F")) && (!data28.equals("F")) && (!data29.equals("F")) && (!data30.equals("F")) && (!data31.equals("F"))
                ) 
                flag1 = 1;
                
            if (data37 >= 3.00 && 
                    (!data38.equals("F")) && (!data39.equals("F")) && (!data40.equals("F")) && (!data41.equals("F")) && (!data42.equals("F")) && (!data43.equals("F")) && (!data44.equals("F")) && (!data45.equals("F")) && (!data46.equals("F")) && (!data47.equals("F")) && (!data48.equals("F")) && (!data49.equals("F")) &&(!data50.equals("F")) && (!data51.equals("F")) && (!data52.equals("F")) && (!data53.equals("F")) && (!data54.equals("F")) && (!data55.equals("F")) && (!data56.equals("F")) && (!data57.equals("F")) && (!data58.equals("F")) && (!data59.equals("F")) && (!data60.equals("F")) && (!data61.equals("F")) && (!data62.equals("F"))
                ) 
                flag2 = 1;
            
            if(flag1 == 1 && flag2 == 1)
            {
                System.out.println("Team " + i + " is eligible for thesis.");
                
                
                
                XSSFRow row2 = worksheet.createRow(1);
                
                XSSFCell cellAm = row2.createCell(0);
		cellAm.setCellValue(data1);
                XSSFCell cellBm = row2.createCell(1);
		cellBm.setCellValue(data2);
                XSSFCell cellCm = row2.createCell(2);
		cellCm.setCellValue(data32);
                XSSFCell cellDm = row2.createCell(3);
		cellDm.setCellValue(data33);
                
                XSSFCell cellEm = row2.createCell(4);
		cellEm.setCellValue(data7);
                XSSFCell cellFm = row2.createCell(5);
		cellFm.setCellValue(data8);
                XSSFCell cellGm = row2.createCell(6);
		cellGm.setCellValue(data9);
                XSSFCell cellHm = row2.createCell(7);
		cellHm.setCellValue(data10);
                XSSFCell cellIm = row2.createCell(8);
		cellIm.setCellValue(data11);
                XSSFCell cellJm = row2.createCell(9);
		cellJm.setCellValue(data12);
                XSSFCell cellKm = row2.createCell(10);
		cellKm.setCellValue(data13);
                XSSFCell cellLm = row2.createCell(11);
		cellLm.setCellValue(data14);
                XSSFCell cellMm = row2.createCell(12);
		cellMm.setCellValue(data15);
                XSSFCell cellNm = row2.createCell(13);
		cellNm.setCellValue(data16);
                XSSFCell cellOm = row2.createCell(14);
		cellOm.setCellValue(data17);
                XSSFCell cellPm = row2.createCell(15);
		cellPm.setCellValue(data18);
                XSSFCell cellQm = row2.createCell(16);
		cellQm.setCellValue(data19);
                XSSFCell cellRm = row2.createCell(17);
		cellRm.setCellValue(data20);
                XSSFCell cellSm = row2.createCell(18);
		cellSm.setCellValue(data21);
                XSSFCell cellTm = row2.createCell(19);
		cellTm.setCellValue(data22);
                XSSFCell cellUm = row2.createCell(20);
		cellUm.setCellValue(data23);
                XSSFCell cellVm = row2.createCell(21);
		cellVm.setCellValue(data24);
                XSSFCell cellWm = row2.createCell(22);
		cellWm.setCellValue(data25);
                XSSFCell cellXm = row2.createCell(23);
		cellXm.setCellValue(data26);
                XSSFCell cellYm = row2.createCell(24);
		cellYm.setCellValue(data27);
                XSSFCell cellZm = row2.createCell(25);
		cellZm.setCellValue(data28);
                XSSFCell cellAAm = row2.createCell(26);
		cellAAm.setCellValue(data29);
                XSSFCell cellABm = row2.createCell(27);
		cellABm.setCellValue(data30);
                XSSFCell cellACm = row2.createCell(28);
		cellACm.setCellValue(data31);
                
                
                
                XSSFCell cellAEm = row2.createCell(30);
		cellAEm.setCellValue(data38);
                XSSFCell cellAFm = row2.createCell(31);
		cellAFm.setCellValue(data39);
                XSSFCell cellAGm = row2.createCell(32);
		cellAGm.setCellValue(data40);
                XSSFCell cellAHm = row2.createCell(33);
		cellAHm.setCellValue(data41);
                XSSFCell cellAIm = row2.createCell(34);
		cellAIm.setCellValue(data42);
                XSSFCell cellAJm = row2.createCell(35);
		cellAJm.setCellValue(data43);
                XSSFCell cellAKm = row2.createCell(36);
		cellAKm.setCellValue(data44);
                XSSFCell cellALm = row2.createCell(37);
		cellALm.setCellValue(data45);
                XSSFCell cellAMm = row2.createCell(38);
		cellAMm.setCellValue(data46);
                XSSFCell cellANm = row2.createCell(39);
		cellANm.setCellValue(data47);
                XSSFCell cellAOm = row2.createCell(40);
		cellAOm.setCellValue(data48);
                XSSFCell cellAPm = row2.createCell(41);
		cellAPm.setCellValue(data49);
                XSSFCell cellAQm = row2.createCell(42);
		cellAQm.setCellValue(data50);
                XSSFCell cellARm = row2.createCell(43);
		cellARm.setCellValue(data51);
                XSSFCell cellASm = row2.createCell(44);
		cellASm.setCellValue(data52);
                XSSFCell cellATm = row2.createCell(45);
		cellATm.setCellValue(data53);
                XSSFCell cellAUm = row2.createCell(46);
		cellAUm.setCellValue(data54);
                XSSFCell cellAVm = row2.createCell(47);
		cellAVm.setCellValue(data55);
                XSSFCell cellAWm = row2.createCell(48);
		cellAWm.setCellValue(data56);
                XSSFCell cellAXm = row2.createCell(49);
		cellAXm.setCellValue(data57);
                XSSFCell cellAYm = row2.createCell(50);
		cellAYm.setCellValue(data58);
                XSSFCell cellAZm = row2.createCell(51);
		cellAZm.setCellValue(data59);
                XSSFCell cellBAm = row2.createCell(52);
		cellBAm.setCellValue(data60);
                XSSFCell cellBBm = row2.createCell(53);
		cellBBm.setCellValue(data61);
                XSSFCell cellBCm = row2.createCell(54);
		cellBCm.setCellValue(data62);
                
                
                
                
                
                
                workbook.write(fileOut);
		fileOut.flush();
		fileOut.close();
                m++;
            }
            else if(flag1 == 1 && flag2 == 0)
            {
                System.out.println("Team " + i + " is not eligible for thesis.\n Because member 2 Reg. No. " + data33 + " hasn't fulfilled the requirements.");
            }
            else if(flag1 == 0 && flag2 == 1)
            {
                System.out.println("Team " + i + " is not eligible for thesis.\n Because member 1 Reg. No. " + data2 + " hasn't fulfilled the requirements.");
            }
            else if(flag1 == 0 && flag2 == 0)
            {
                System.out.println("Team " + i + " is not eligible for thesis.\n Because none of them (Reg. No. " + data2 + " and " + data33 + ") fulfilled the requirements.");
            }
        }
 
    } 
    catch (Exception e)
    {
        System.out.println(e.getMessage());
    }
    
}
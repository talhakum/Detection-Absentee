/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package annualleave;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.time.Month;
import java.util.Calendar;
import java.util.Iterator;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author talha
 */
public class PersonelTara {
    static String tarih;
    static String ad;
    static String GETIR[]=new String[5000];
    static int z=0;
    static int i=0;
    static long toplam=0;
    static String ayYil;
    static String deneme;
    static int test=0;
    static int c=0;
    static String yeniAd[]=new String[100];
    static boolean isBuilt=true;
    static int yil,ay;
    static public void Build() {
        isBuilt=false;
    }
    static public void Detection(int ilk, int son,String URL) throws FileNotFoundException, IOException {
        char gun[]=new char[100];
        int gunler[]=new int[32];
        for(int j=ilk;j<=son;j++) {
            gunler[j]=j;
        }
        String oncekiAd;
        // C:\\Users\\talha\\Documents\\NetBeansProjects\\AnnualLeave\\src\\annualleave\\Mayıs 23.xlsx
            String excelFilePath =URL;
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int sonuncuIndex=firstSheet.getLastRowNum();
        Iterator<Row> iterator = firstSheet.iterator();
        
        
        while(iterator.hasNext()) {

            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            
            if(cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getRowIndex()>=6 ) {
                                   oncekiAd=ad;
                cell=cellIterator.next();
                cell=cellIterator.next();
                cell=cellIterator.next();
                ad=cell.getStringCellValue();
                cell=cellIterator.next();
                cell=cellIterator.next();
                tarih=cell.getStringCellValue();
                if(ad!=oncekiAd &&  i!=0  && !(oncekiAd.equals("Personel Adı Soyadı")) && 
                        !(oncekiAd.isEmpty()) && !(ad.isEmpty()) && !(ad.equals("Personel Adı Soyadı")) ||
                        cell.getRowIndex()==sonuncuIndex ) {
             for(int j=ilk;j<=son;j++) {
                   if(gunler[j]!=0) {
               GETIR[z]=oncekiAd+" "+gunler[j];
                   z++;
                   test=1;
                   }
                   if(isBuilt) {
                       int left=tarih.indexOf(".");
                       int right=tarih.lastIndexOf(".");
                       String sub=tarih.substring(left+1,right);
                       ay=Integer.parseInt(sub);
                       
                       int left2=tarih.lastIndexOf(".");
                       int right2=tarih.lastIndexOf("");
                       String sub2=tarih.substring(left2+1,right2);
                       yil=Integer.parseInt(sub2);
                   }
                   Build();
           }
             if(test==1) {
                 yeniAd[c]=oncekiAd;
                  for(int j=ilk;j<=son;j++) {
                   if(gunler[j]!=0) {
                    Calendar date = Calendar.getInstance();
                date.set(yil,ay-1,gunler[j]);
                if(date.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || date.get(Calendar.DAY_OF_WEEK) == 
                        Calendar.SUNDAY) {
                }
                else {
                       yeniAd[c]+=" "+gunler[j];
                }
                   }
           }
                 c++;
                 test=0;
             }
               for(int j=ilk;j<=son;j++) {
                     gunler[j]=j;
        }
                }
                if( !(cell.getStringCellValue().isEmpty()) && !(ad.equals("Personel Adı Soyadı")) ) {
               int left = tarih.indexOf(0);
               int right = tarih.indexOf(".");
               String sub = tarih.substring(left+1, right);
               gunler[Integer.parseInt(sub)]=0;
               i++;
                }
                }
            }
    }
        }
    
    public static void main(String[] args) throws FileNotFoundException, IOException {
        Detection(2, 23, "C:\\Users\\talha\\Documents\\NetBeansProjects\\AnnualLeave\\src\\annualleave\\Mayıs 23.xlsx");
        for(int k=0; k<c; k++) {
            System.out.println(yeniAd[k]+" MAYIS 2016 tarihlerinde şirkete giriş yapmamıştır.");
        }
    }
        }
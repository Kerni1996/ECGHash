import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.*;

public class main {
    private static final String ExccelPath = "./Report.xlsx";
    private static final String ECGPath = "./ecg-liste.hash";
    private static final String columnEmail = "E-Mail";
    private static String ECGHash = "";
    public static void main(String[] args) throws IOException, InvalidFormatException, NoSuchAlgorithmException, EmailColumnNotFoundException {
        System.out.println("Read by File: ");
        countASCIILines(new File(ECGPath));
        System.out.println("removed Adresses: " + alterExcel());
        System.out.println();

        byte[] iso88591Data = "karl.testinger@firma.at".toLowerCase().getBytes("ISO-8859-1");
        System.out.println("ISO karl.testinger@firma.at");
        System.out.println(ECGHash.contains(toSHA1(iso88591Data)));
        System.out.println();

        iso88591Data="max.mustermann@home.at".getBytes("ISO-8859-1");
        System.out.println("ISO max.mustermann@home.at");
        System.out.println(ECGHash.contains(toSHA1(iso88591Data)));
        System.out.println();

        iso88591Data="@home.lan".getBytes("ISO-8859-1");
        System.out.println("ISO @home.lan");
        System.out.println(ECGHash.contains(toSHA1(iso88591Data)));
        System.out.println();

        iso88591Data="@firma.lan".getBytes("ISO-8859-1");
        System.out.println("ISO @firma.lan");
        System.out.println(ECGHash.contains(toSHA1(iso88591Data)));
        System.out.println();



    }



    public static String toSHA1(byte[] convertme) {
        MessageDigest md = null;
        try {
            md = MessageDigest.getInstance("SHA-1");
        }
        catch(NoSuchAlgorithmException e) {
            e.printStackTrace();
        }
        return new String(md.digest(convertme));
    }


    public static boolean countASCIILines(File f) throws IOException {
        BufferedReader br = new BufferedReader(new InputStreamReader(
                new FileInputStream(f)));
        try {
            int count = 0;
            String line;
            while ((line = br.readLine()) != null) {
                ECGHash = ECGHash+line;
                count++;
            }
            return true;
        } finally {
            br.close();
        }
    }







    private  static LinkedList<String> alterExcel() throws IOException, InvalidFormatException, EmailColumnNotFoundException {
        LinkedList<String> removedAddresses = new LinkedList<String>();
        Workbook workbook = WorkbookFactory.create(new File(ExccelPath));
        Sheet sheet = workbook.getSheetAt(0);
        Row firstRow = sheet.getRow(0);
        int indexEmailColumn = -1;
        for (int i = 0; i<firstRow.getLastCellNum(); i++){
            if (firstRow.getCell(i).getStringCellValue().equals(columnEmail)){
                indexEmailColumn = i;
                break;
            }
        }

        if (indexEmailColumn == -1){

           throw new EmailColumnNotFoundException();
        }

        for (int i = 1; i<sheet.getLastRowNum(); i++){
            Row row = sheet.getRow(i);


            //System.out.println(sheet.getRow(i));
            if (row!=null) {
                String email = sheet.getRow(i).getCell(indexEmailColumn).getStringCellValue();
                //convert email String to lower case
                email = email.toLowerCase();

                //check also if domain is in ECg List
                String domain = "@"+email.split("@")[1];

                byte[] iso88591Mail = email.getBytes("ISO-8859-1");
                byte[] iso88591Domain = domain.getBytes("ISO-8859-1");

                //apply hash
                String sha1Mail = toSHA1(iso88591Mail);
                String sha1Domain = toSHA1(iso88591Domain);

                if (ECGHash.contains(sha1Mail)||ECGHash.contains(sha1Domain)){
                    removedAddresses.add(email);
                }



            }

        }
        workbook.close();
        return removedAddresses;
    }

}

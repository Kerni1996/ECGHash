import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import java.io.*;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.*;

public class main {

    private static final String columnEmail = "E-Mail";
    private static String ECGHash = "";
    public static void main(String[] args) throws IOException, InvalidFormatException, NoSuchAlgorithmException, EmailColumnNotFoundException {

        readHashFile(selectHash());
        alterExcel(selectExcel());

    }

    private static File selectHash(){
        JFileChooser chooser= new JFileChooser();
        chooser.setFileFilter(new FileFilter() {
            @Override
            public boolean accept(File f) {
                if (f.isDirectory()){
                    return true;
                } else {
                    String filename= f.getName().toLowerCase();
                    return filename.endsWith(".hash");
                }
            }

            @Override
            public String getDescription() {
                return "Hash files (*.hash)";
            }
        });

        int selection = chooser.showDialog(null,"Please select Hash file with email-adresses");
        if (selection == JFileChooser.APPROVE_OPTION && chooser.getSelectedFile().getAbsolutePath().endsWith(".hash")){
            return chooser.getSelectedFile().getAbsoluteFile();
        }else return selectHash();
    }

    private static File selectExcel(){
        JFileChooser chooser= new JFileChooser();
        chooser.setFileFilter(new FileFilter() {
            @Override
            public boolean accept(File f) {
                if (f.isDirectory()){
                    return true;
                } else {
                    String filename= f.getName().toLowerCase();
                    return filename.endsWith(".xlsx");
                }
            }

            @Override
            public String getDescription() {
                return "Excel files (*.xlsx)";
            }
        });

        int selection = chooser.showDialog(null,"Please select Excel file with email-adresses");
        if (selection == JFileChooser.APPROVE_OPTION && chooser.getSelectedFile().getAbsolutePath().endsWith(".xlsx")){
            return chooser.getSelectedFile().getAbsoluteFile();
        }else return selectExcel();
    }



    private static String toSHA1(byte[] convertme) {
        MessageDigest md = null;
        try {
            md = MessageDigest.getInstance("SHA-1");
        }
        catch(NoSuchAlgorithmException e) {
            e.printStackTrace();
        }
        return new String(md.digest(convertme));
    }


    private static boolean readHashFile(File f) throws IOException {
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



    private  static LinkedList<String> alterExcel(File file) throws IOException, InvalidFormatException, EmailColumnNotFoundException {


        LinkedList<String> removedAddresses = new LinkedList<String>();
        Workbook workbook = WorkbookFactory.create(file);
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
                    sheet.removeRow(sheet.getRow(i));
                    //sheet.getRow(i).createCell(indexEmailColumn+1).setCellValue(true);
                    removedAddresses.add(email);
                }
                //else sheet.getRow(i).createCell(indexEmailColumn+1).setCellValue(false);



            }

        }
        FileOutputStream out = new FileOutputStream(file.getAbsolutePath().replace(".xlsx","-robinsonChecked.xlsx"));

        workbook.write(out);
        workbook.close();


        JOptionPane.showMessageDialog(null,"The following recipients were removed from the Excel file (" + file.getAbsolutePath() + "):\n" + removedAddresses, "Checked with Robinson",JOptionPane.INFORMATION_MESSAGE);

        return removedAddresses;


    }

}

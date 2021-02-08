import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        ArrayList<Object> list = new ArrayList<Object>( );
        Object array[] = new Object[21];


        String fileName = "//Desktop//Copy of Network Delay Disruption Particulars - KF(112).xlsx";
        File file = new File(System.getProperty("user.home"), fileName);
        try {

            Connection connection = DBConnection.getDbConnection( ).getConnection( );
            String sql = "INSERT INTO execlData(OriginalStartDate,OriginalTaskFinishDay,StartOfDelay ,EndOfDelay,CauseOfDetay,EffectofDelay,EngineerCommentCause,EngineerCommentEffect,CommercialComment,TaskName,EventReference,EvidenceIssued,ResourceNames,ResourceGrade,NoOfMen,NoOfHours,NoOfMinute,TotalOdDuration,TotalCost) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
            connection.setAutoCommit(false);

            PreparedStatement pstm = connection.prepareStatement(sql);


            Workbook workbook = WorkbookFactory.create(file);
            System.out.println(workbook);

            Iterator<Sheet> sheetIterator = workbook.sheetIterator( );
            while (sheetIterator.hasNext( )) {
                Sheet sheet = sheetIterator.next( );
                System.out.println(sheet.getSheetName( ));
            }
            Sheet sheetAt = workbook.getSheetAt(0);

            DataFormatter dataFormatter = new DataFormatter( );

            Iterator<Row> rowIterator = sheetAt.rowIterator( );
            int rowIndex = 0;
            int cellIndex = 0;

            while (rowIterator.hasNext( )) {
                Row row = rowIterator.next( );
                rowIndex++;
                Iterator<Cell> cellIterator = row.cellIterator( );

                while (cellIterator.hasNext( )) {
                    cellIndex++;
                    Cell cell = cellIterator.next( );
//                    String s = dataFormatter.formatCellValue(cell);
//                    System.out.print(s+"\t");
                    int columnIndex = cell.getColumnIndex( );


                    if (rowIndex > 6) {

                        if (columnIndex > 3 && columnIndex < 23) {
                            System.out.print("[" + columnIndex + "]");
                            String stringCellValue="";
                            String s="";
                            String d="";



                                    CellType type = CellType.forInt(cell.getCellType( ));
                            CellStyle cellStyle = null;


                            if (type == CellType.FORMULA) {
                                type = cell.getCachedFormulaResultTypeEnum( );

                            }
                            if (type == CellType.NUMERIC) {
                                 cellStyle= cell.getCellStyle( );

                                if (cellStyle != null && cellStyle.getDataFormatString( ) != null) {
                                   s = dataFormatter.formatRawCellContents(cell.getNumericCellValue( ), cellStyle.getDataFormat( ), cellStyle.getDataFormatString( ));
                                    System.out.print(s + "\t");
                                }
                            }
                            if (type == CellType.STRING) {
                                stringCellValue = cell.getStringCellValue( );
                                System.out.print(stringCellValue + "\t");
                            }

                            switch (columnIndex){
                                case 4:
                                    pstm.setString(1,s);
                                case 5:
                                    pstm.setString(2,s);
                                case 6:pstm.setString(3,s);
                                case 7:pstm.setString(4,s);
                                case 8:
                                    pstm.setString(5,stringCellValue);
                                case 9:
                                    pstm.setString(6,stringCellValue);
                                case 10:
                                    pstm.setString(7,stringCellValue);
                                case 11:
                                    pstm.setString(8,stringCellValue);
                                case 12:
                                    pstm.setString(9,d);
                                case 13:
                                    pstm.setString(10,stringCellValue);
                                case 15:
                                    pstm.setString(12,stringCellValue);
                                case 16:
                                    pstm.setString(13,stringCellValue);
                                case 17:
                                    pstm.setString(14,d);
                                case 14:pstm.setString(11,s);
                                case 18:pstm.setString(15,s);
                                case 19:pstm.setString(16,s);
                                case 20:pstm.setString(17,s);
                                case 21:pstm.setString(18,s);
                                case 22:pstm.setString(19,s);
                                default:



                            }

                            connection.commit();
                            pstm.executeUpdate();

                        }



                    }
                    System.out.println( );

//                for (Row r: sheetAt) {
//                    for(Cell cell: row) {
//                        System.out.println(cell);
//                    }
//                    System.out.println();
//                }

                }

            }

                workbook.close( );

            } catch(IOException e){

            } catch(InvalidFormatException e){

            } catch(SQLException e){

            } catch(ClassNotFoundException e){

            }
        }
    }


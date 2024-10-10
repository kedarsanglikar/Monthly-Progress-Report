/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package javaapplication21;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.sql.DatabaseMetaData;
import java.util.StringTokenizer;
import javax.swing.DefaultComboBoxModel;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ComboBoxModel;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdesktop.swingx.autocomplete.AutoCompleteDecorator;

/**
 *
 * @author Dell
 */
public class AddData extends javax.swing.JFrame {

    SimpleDateFormat formatter = new SimpleDateFormat("dd-MMM-y");

    SimpleDateFormat toDataBaseDate = new SimpleDateFormat("yyyy-MM-dd");
    Config config;
    DatabaseMetaData metaData;
    ResultSet rs, rs2;
    String tables = "";
    String year;

    String bandkamVibhag;
    String taluka;
    String tableName;
    String kamacheNaavTxtStr;
    String prashashkiyaDinankStr;
    String manyataRakkamStr;
    String tantrikDinankStr;
    String tantrikManyatRakkamStr, maktedaracheNaavTxtStr, kamacheAdeshDinankStr, nivedaRakkamStr, gstTxtStr, akunTxtStr, kamachiMudatMahineStr, praptSunStr, kharchWthGstStr, sheraStr;
    String workCompleteDateStr;
    ArrayList<String> TABLE_COLUMNS, TABLE_COLUMNS2;
    DefaultTableModel model, model2;
    Statement stmt, stmt2;
    int ROW_HEIGHT = 30;
    String selectedTableName;
    long nivedaRakkamLong, gstTxtLong, nivedaRakkamLong1, gstTxtLong1;
    String location = null;
    String filename;

    /**
     * Creates new form AddData
     */
    public AddData() {
        initComponents();
        config = new Config();
        getTables();
        disableUpdateFields(false);
        System.out.println(sheraComboBox.getItemCount());
        System.out.println(sheraComboBox1.getItemCount());

        AutoCompleteDecorator.decorate(sheraComboBox);
        AutoCompleteDecorator.decorate(sheraComboBox1);
        AutoCompleteDecorator.decorate(yojnaComboBox);
        AutoCompleteDecorator.decorate(kamacheNaavComboBox);
        AutoCompleteDecorator.decorate(updateyojnaComboBox);

        AutoCompleteDecorator.decorate(vibhaagComboBox);
        AutoCompleteDecorator.decorate(vibhaagComboBox1);
        AutoCompleteDecorator.decorate(talukaComboBox);

        AutoCompleteDecorator.decorate(talukaComboBox1);
        AutoCompleteDecorator.decorate(yearComboBox);
        AutoCompleteDecorator.decorate(yearComboBox1);

    }

    public void getTables() {

        try {
            DatabaseMetaData metaData = config.conn.getMetaData();

            ResultSet tables = metaData.getTables(null, null, "%", new String[]{"TABLE"});
            ArrayList<String> tableNames = new ArrayList<>();

            while (tables.next()) {
                tableNames.add(tables.getString("TABLE_NAME"));
            }

            // Convert ArrayList to array
            String[] tableNamesArray = tableNames.toArray(new String[0]);

            // Print table names
            for (String tableName : tableNamesArray) {
                System.out.println(tableName);
            }
            yojnaComboBox.setModel(new DefaultComboBoxModel(tableNamesArray));
            updateyojnaComboBox.setModel(new DefaultComboBoxModel(tableNamesArray));
            displayYojnaComboBox.setModel(new DefaultComboBoxModel(tableNamesArray));
        } catch (SQLException e) {
            e.printStackTrace();
        }

    }

    private void clearAllFields() {
        kamacheNaavTxt.setText("");
        prashashkiyaDinank.setDate(null);
        manyataRakkam.setText("");
        tantrikDinank.setDate(null);
        tantrikManyatRakkam.setText("");
        maktedaracheNaavTxt.setText("");
        kamacheAdeshDinank.setDate(null);
        nivedaRakkam.setText("");
        gstTxt.setText("");
        akunTxt.setText("");
        kamachiMudatMahineDate.setDate(null);
        praptSun.setText("");
        kharchWthGst.setText("");
        workCompleteDate.setDate(null);

//        kamacheNaavTxt.setText("");
        prashashkiyaDinank1.setDate(null);
        manyataRakkam1.setText("");
        tantrikDinank1.setDate(null);
        tantrikManyatRakkam1.setText("");
        maktedaracheNaavTxt1.setText("");
        kamacheAdeshDinank1.setDate(null);
        nivedaRakkam1.setText("");
        gstTxt1.setText("");
        akunTxt1.setText("");
        kamachiMudatMahineDate1.setDate(null);
        praptSunUpd.setText("");
        kharchWthGstUpd.setText("");
        workCompleteDate1.setDate(null);

    }

    private void disableUpdateFields(boolean flag) {
        yearComboBox1.setEnabled(flag);
        vibhaagComboBox1.setEnabled(flag);
        talukaComboBox1.setEnabled(flag);
        prashashkiyaDinank1.setEnabled(flag);
        manyataRakkam1.setEnabled(flag);
        tantrikDinank1.setEnabled(flag);
        tantrikManyatRakkam1.setEnabled(flag);
        maktedaracheNaavTxt1.setEnabled(flag);
        kamacheAdeshDinank1.setEnabled(flag);
        nivedaRakkam1.setEnabled(flag);
        gstTxt1.setEnabled(flag);
//        akunTxt1.setEnabled(flag);
        kamachiMudatMahineDate1.setEnabled(flag);
        praptSunUpd.setEnabled(flag);
        kharchWthGstUpd.setEnabled(flag);
        workCompleteDate1.setEnabled(flag);
        sheraComboBox1.setEnabled(flag);

        yearComboBox1.setEditable(flag);
        vibhaagComboBox1.setEditable(flag);
        talukaComboBox1.setEditable(flag);
        manyataRakkam1.setEditable(flag);
        tantrikManyatRakkam1.setEditable(flag);
        maktedaracheNaavTxt1.setEditable(flag);
        nivedaRakkam1.setEditable(flag);
        gstTxt1.setEditable(flag);
//        akunTxt1.setEditable(flag);

        praptSunUpd.setEditable(flag);
        kharchWthGstUpd.setEditable(flag);
        sheraComboBox1.setEditable(flag);

    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPanel2 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        yojnaComboBox = new javax.swing.JComboBox<>();
        jLabel2 = new javax.swing.JLabel();
        kamacheNaavTxt = new javax.swing.JTextField();
        prashashkiyaDinank = new com.toedter.calendar.JDateChooser();
        jLabel3 = new javax.swing.JLabel();
        manyataRakkam = new javax.swing.JTextField();
        jLabel4 = new javax.swing.JLabel();
        tantrikDinank = new com.toedter.calendar.JDateChooser();
        jLabel5 = new javax.swing.JLabel();
        tantrikManyatRakkam = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        saveBtn = new javax.swing.JButton();
        jLabel7 = new javax.swing.JLabel();
        maktedaracheNaavTxt = new javax.swing.JTextField();
        jLabel8 = new javax.swing.JLabel();
        kamacheAdeshDinank = new com.toedter.calendar.JDateChooser();
        jLabel9 = new javax.swing.JLabel();
        nivedaRakkam = new javax.swing.JTextField();
        jLabel10 = new javax.swing.JLabel();
        gstTxt = new javax.swing.JTextField();
        jLabel11 = new javax.swing.JLabel();
        akunTxt = new javax.swing.JTextField();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        praptSun = new javax.swing.JTextField();
        jLabel15 = new javax.swing.JLabel();
        kharchWthGst = new javax.swing.JTextField();
        jLabel17 = new javax.swing.JLabel();
        jLabel36 = new javax.swing.JLabel();
        yearComboBox = new javax.swing.JComboBox<>();
        jLabel14 = new javax.swing.JLabel();
        vibhaagComboBox = new javax.swing.JComboBox<>();
        jLabel16 = new javax.swing.JLabel();
        talukaComboBox = new javax.swing.JComboBox<>();
        sheraComboBox = new javax.swing.JComboBox<>();
        jLabel38 = new javax.swing.JLabel();
        workCompleteDate = new com.toedter.calendar.JDateChooser();
        kamachiMudatMahineDate = new com.toedter.calendar.JDateChooser();
        jButton1 = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jLabel19 = new javax.swing.JLabel();
        updateyojnaComboBox = new javax.swing.JComboBox<>();
        jLabel18 = new javax.swing.JLabel();
        kamacheNaavComboBox = new javax.swing.JComboBox<>();
        jLabel21 = new javax.swing.JLabel();
        prashashkiyaDinank1 = new com.toedter.calendar.JDateChooser();
        jLabel22 = new javax.swing.JLabel();
        manyataRakkam1 = new javax.swing.JTextField();
        jLabel23 = new javax.swing.JLabel();
        tantrikDinank1 = new com.toedter.calendar.JDateChooser();
        jLabel24 = new javax.swing.JLabel();
        tantrikManyatRakkam1 = new javax.swing.JTextField();
        jLabel25 = new javax.swing.JLabel();
        maktedaracheNaavTxt1 = new javax.swing.JTextField();
        jLabel26 = new javax.swing.JLabel();
        kamacheAdeshDinank1 = new com.toedter.calendar.JDateChooser();
        jLabel27 = new javax.swing.JLabel();
        nivedaRakkam1 = new javax.swing.JTextField();
        jLabel28 = new javax.swing.JLabel();
        gstTxt1 = new javax.swing.JTextField();
        jLabel29 = new javax.swing.JLabel();
        akunTxt1 = new javax.swing.JTextField();
        jLabel30 = new javax.swing.JLabel();
        jLabel31 = new javax.swing.JLabel();
        praptSunUpd = new javax.swing.JTextField();
        jLabel34 = new javax.swing.JLabel();
        kharchWthGstUpd = new javax.swing.JTextField();
        jLabel35 = new javax.swing.JLabel();
        updateBtn = new javax.swing.JButton();
        editBtn = new javax.swing.JButton();
        sheraComboBox1 = new javax.swing.JComboBox<>();
        jLabel32 = new javax.swing.JLabel();
        yearComboBox1 = new javax.swing.JComboBox<>();
        jLabel33 = new javax.swing.JLabel();
        vibhaagComboBox1 = new javax.swing.JComboBox<>();
        jLabel37 = new javax.swing.JLabel();
        talukaComboBox1 = new javax.swing.JComboBox<>();
        jLabel39 = new javax.swing.JLabel();
        workCompleteDate1 = new com.toedter.calendar.JDateChooser();
        kamachiMudatMahineDate1 = new com.toedter.calendar.JDateChooser();
        deleteBtn = new javax.swing.JButton();
        jPanel4 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        yojnaNorthTable = new javax.swing.JTable();
        displayYojnaComboBox = new javax.swing.JComboBox<>();
        jLabel20 = new javax.swing.JLabel();
        showDataBtn = new javax.swing.JButton();
        jPanel5 = new javax.swing.JPanel();
        exportToExcel = new javax.swing.JButton();
        closeBtn = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        abstractTable1 = new javax.swing.JTable();
        jPanel6 = new javax.swing.JPanel();
        exportToExcel1 = new javax.swing.JButton();
        closeBtn1 = new javax.swing.JButton();
        jScrollPane4 = new javax.swing.JScrollPane();
        abstractTable2 = new javax.swing.JTable();
        jPanel7 = new javax.swing.JPanel();
        exportToExcel2 = new javax.swing.JButton();
        closeBtn2 = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        yojnaSouthTable = new javax.swing.JTable();
        jPanel8 = new javax.swing.JPanel();
        exportToExcel3 = new javax.swing.JButton();
        closeBtn3 = new javax.swing.JButton();
        yearChooseDisplay = new javax.swing.JComboBox<>();
        allDataButton = new javax.swing.JButton();
        showMaktedarForm = new javax.swing.JButton();
        jPanel9 = new javax.swing.JPanel();
        backupBtn = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setTitle("मासिक प्रगती अहवाल");
        setResizable(false);

        jPanel1.setBackground(new java.awt.Color(204, 255, 255));

        jTabbedPane1.setBackground(new java.awt.Color(255, 255, 255));
        jTabbedPane1.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jPanel2.setBackground(new java.awt.Color(204, 255, 255));
        jPanel2.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel1.setText("योजना निवडा :");
        jLabel1.setToolTipText("");

        yojnaComboBox.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        yojnaComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "योजना निवडा" }));

        jLabel2.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel2.setText("कामाचे नाव : ");

        kamacheNaavTxt.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel3.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel3.setText("प्रशासकीय मान्यता दिनाक : ");

        manyataRakkam.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel4.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel4.setText("प्रशासकीय मान्यता रक्कम :");

        jLabel5.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel5.setText("तांत्रीक मान्यता दिनांक :");

        tantrikManyatRakkam.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel6.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel6.setText("तांत्रीक मान्यता रक्कम :");

        saveBtn.setFont(new java.awt.Font("Mangal", 1, 12)); // NOI18N
        saveBtn.setText("साठीवणे");
        saveBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveBtnActionPerformed(evt);
            }
        });

        jLabel7.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel7.setText("मक्तेदाराचे नांव :");

        maktedaracheNaavTxt.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel8.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        jLabel8.setText("कामाचा आदेश व दिनांक : ");

        jLabel9.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel9.setText("निविदा स्विकृती रक्कम :");

        nivedaRakkam.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel10.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel10.setText("जी एस टी :");

        gstTxt.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel11.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel11.setText("एकुण :");

        akunTxt.setEditable(false);
        akunTxt.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        akunTxt.setEnabled(false);

        jLabel12.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel12.setText("कामाची मुदत महिने :");

        jLabel13.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel13.setText("प्राप्त निधी या वर्षा मध्ये :");

        praptSun.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        praptSun.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                praptSunFocusGained(evt);
            }
        });

        jLabel15.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel15.setText("या वर्षा मधील खर्च जी एस टी सह :");

        kharchWthGst.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel17.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel17.setText("शेरा :");

        jLabel36.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel36.setText("वर्ष निवडा :");

        yearComboBox.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N
        yearComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021-2022", "2022-2023", "2023-2024", "2024-2025", "2025-2026" }));

        jLabel14.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel14.setText("बांधकाम विभाग :");

        vibhaagComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        vibhaagComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "विभाग निवडा", "उत्तर विभाग", "दक्षिण विभाग" }));
        vibhaagComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                vibhaagComboBoxActionPerformed(evt);
            }
        });

        jLabel16.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel16.setText("तालुका :");

        talukaComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        talukaComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "तालुका निवडा" }));

        sheraComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        sheraComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "काम रद्द प्रस्तावित ", "काम रद्द", "काम पूर्ण", "जागा अडचण (कोर्ट केस)", "(वनविभाग) जागा अडचण", "स्थानिक जागा अडचण", "पाय खुदाई प्रगतीत", "जोते स्तर (फ्लिंथ)", "चौकट स्तर", "छत/ स्लॅब स्तर", "प्लास्टर", "फरशी", "रंगकाम", "भौतिकदृष्ट्या काम पूर्ण", "सुरू नाही", "जी -१ प्रगतीत", "जी -२ प्रगतीत ", "जी -१ खंडिकरण पूर्ण", "जी -२ खंडिकरण पूर्ण", "एम पी एम प्रगतीत", "एम पी एम पूर्ण", "कार्पेट प्रगतीत", "कार्पेट पूर्ण", "सिळकोट प्रगतीत", "सिळकोट पूर्ण", "मातीकाम प्रगतीत", "मातीकाम पूर्ण", "मोरी बांधकाम प्रगतीत", "मोरी बांधकाम पूर्ण", "पाईप बसविणे प्रगतीत", "पाईप बसविणे पूर्ण", "स्लॅब ड्रेन प्रगतीत", "स्लॅब ड्रेन पूर्ण" }));

        jLabel38.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel38.setText("काम पूर्ण झालेल्याचे तारिक :");

        workCompleteDate.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N

        kamachiMudatMahineDate.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N
        kamachiMudatMahineDate.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                kamachiMudatMahineDateFocusGained(evt);
            }
        });
        kamachiMudatMahineDate.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                kamachiMudatMahineDateMouseClicked(evt);
            }
        });

        jButton1.setFont(new java.awt.Font("Mangal", 1, 13)); // NOI18N
        jButton1.setText("बंद करा");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                        .addGap(20, 20, 20)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 143, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 178, Short.MAX_VALUE)
                                        .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                    .addComponent(jLabel3)))
                            .addComponent(jLabel36, javax.swing.GroupLayout.PREFERRED_SIZE, 81, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 156, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel14, javax.swing.GroupLayout.PREFERRED_SIZE, 167, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel16, javax.swing.GroupLayout.PREFERRED_SIZE, 141, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(manyataRakkam, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(tantrikManyatRakkam, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(tantrikDinank, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(yojnaComboBox, 0, 267, Short.MAX_VALUE)
                            .addComponent(yearComboBox, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(prashashkiyaDinank, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(kamacheNaavTxt, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(vibhaagComboBox, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(talukaComboBox, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addGap(97, 97, 97)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 206, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 178, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel15)
                            .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 178, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel38, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(maktedaracheNaavTxt, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(kamacheAdeshDinank, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(kharchWthGst, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(praptSun, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(akunTxt, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(gstTxt, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(nivedaRakkam, javax.swing.GroupLayout.DEFAULT_SIZE, 267, Short.MAX_VALUE)
                            .addComponent(workCompleteDate, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(kamachiMudatMahineDate, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(302, 302, 302)
                        .addComponent(jLabel17, javax.swing.GroupLayout.PREFERRED_SIZE, 59, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(24, 24, 24)
                        .addComponent(sheraComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 255, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(328, 328, 328)
                        .addComponent(saveBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(79, 79, 79)
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel36)
                    .addComponent(yearComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel7)
                    .addComponent(maktedaracheNaavTxt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(kamacheAdeshDinank, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel14)
                        .addComponent(vibhaagComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel8)))
                .addGap(21, 21, 21)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel16)
                            .addComponent(talukaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel9)
                            .addComponent(nivedaRakkam, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(23, 23, 23)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1)
                            .addComponent(yojnaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel10)
                            .addComponent(gstTxt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(23, 23, 23)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel2)
                            .addComponent(kamacheNaavTxt, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel11)
                            .addComponent(akunTxt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jLabel3)
                                .addComponent(prashashkiyaDinank, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel12))
                            .addComponent(kamachiMudatMahineDate, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGap(71, 71, 71)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(jLabel5)
                                    .addComponent(tantrikDinank, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel4)
                                    .addComponent(manyataRakkam, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel13)
                                    .addComponent(praptSun, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel15)
                                    .addComponent(kharchWthGst, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel6)
                            .addComponent(tantrikManyatRakkam, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel38)))
                    .addComponent(workCompleteDate, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(43, 43, 43)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(sheraComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel17))
                .addGap(33, 33, 33)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(saveBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(108, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("माहिती भरणे", jPanel2);

        jPanel3.setBackground(new java.awt.Color(255, 204, 204));
        jPanel3.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel19.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel19.setText("योजना निवडा :");

        updateyojnaComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        updateyojnaComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "योजना निवडा" }));
        updateyojnaComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                updateyojnaComboBoxActionPerformed(evt);
            }
        });

        jLabel18.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel18.setText("कामाचे नाव : ");

        kamacheNaavComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        kamacheNaavComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "काम निवडा" }));
        kamacheNaavComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                kamacheNaavComboBoxActionPerformed(evt);
            }
        });

        jLabel21.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel21.setText("प्रशासकीय मान्यता दिनाक : ");

        jLabel22.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel22.setText("प्रशासकीय मान्यता रक्कम :");

        manyataRakkam1.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel23.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel23.setText("तांत्रीक मान्यता दिनांक :");

        jLabel24.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel24.setText("तांत्रीक मान्यता रक्कम :");

        tantrikManyatRakkam1.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        jLabel25.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel25.setText("मक्तेदाराचे नांव :");

        maktedaracheNaavTxt1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel26.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        jLabel26.setText("कामाचा आदेश व दिनांक : ");

        jLabel27.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel27.setText("निविदा स्विकृती रक्कम :");

        nivedaRakkam1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel28.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel28.setText("जी एस टी :");

        gstTxt1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel29.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel29.setText("एकुण :");

        akunTxt1.setEditable(false);
        akunTxt1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        akunTxt1.setEnabled(false);

        jLabel30.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel30.setText("कामाची मुदत महिने :");

        jLabel31.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel31.setText("प्राप्त निधी या वर्षा मध्ये :");

        praptSunUpd.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        praptSunUpd.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                praptSunUpdFocusGained(evt);
            }
        });

        jLabel34.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel34.setText("या वर्षा मधील खर्च जी एस टी सह :");

        kharchWthGstUpd.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N

        jLabel35.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel35.setText("शेरा :");

        updateBtn.setFont(new java.awt.Font("Mangal", 1, 14)); // NOI18N
        updateBtn.setText("माहिती अपडेट करा");
        updateBtn.setEnabled(false);
        updateBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                updateBtnActionPerformed(evt);
            }
        });

        editBtn.setFont(new java.awt.Font("Mangal", 1, 14)); // NOI18N
        editBtn.setText("एनेबल टू एडिट");
        editBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                editBtnActionPerformed(evt);
            }
        });

        sheraComboBox1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        sheraComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "काम रद्द प्रस्तावित ", "काम रद्द", "काम पूर्ण", "जागा अडचण (कोर्ट केस)", "(वनविभाग) जागा अडचण", "स्थानिक जागा अडचण", "पाय खुदाई प्रगतीत", "जोते स्तर (फ्लिंथ)", "चौकट स्तर", "छत/ स्लॅब स्तर", "प्लास्टर", "फरशी", "रंगकाम", "भौतिकदृष्ट्या काम पूर्ण", "सुरू नाही", "जी -१ प्रगतीत", "जी -२ प्रगतीत ", "जी -१ खंडिकरण पूर्ण", "जी -२ खंडिकरण पूर्ण", "एम पी एम प्रगतीत", "एम पी एम पूर्ण", "कार्पेट प्रगतीत", "कार्पेट पूर्ण", "सिळकोट प्रगतीत", "सिळकोट पूर्ण", "मातीकाम प्रगतीत", "मातीकाम पूर्ण", "मोरी बांधकाम प्रगतीत", "मोरी बांधकाम पूर्ण", "पाईप बसविणे प्रगतीत", "पाईप बसविणे पूर्ण", "स्लॅब ड्रेन प्रगतीत", "स्लॅब ड्रेन पूर्ण" }));

        jLabel32.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel32.setText("वर्ष निवडा :");

        yearComboBox1.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N
        yearComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021-2022", "2022-2023", "2023-2024", "2024-2025", "2025-2026" }));

        jLabel33.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel33.setText("बांधकाम विभाग :");

        vibhaagComboBox1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        vibhaagComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "विभाग निवडा", "उत्तर विभाग", "दक्षिण विभाग" }));
        vibhaagComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                vibhaagComboBox1ActionPerformed(evt);
            }
        });

        jLabel37.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel37.setText("तालुका :");

        talukaComboBox1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        talukaComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "तालुका निवडा" }));

        jLabel39.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel39.setText("काम पूर्ण झालेल्याचे तारिक :");

        workCompleteDate1.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N

        kamachiMudatMahineDate1.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N

        deleteBtn.setFont(new java.awt.Font("Mangal", 1, 14)); // NOI18N
        deleteBtn.setText("हे काम डिलीट करा");
        deleteBtn.setToolTipText("");
        deleteBtn.setEnabled(false);
        deleteBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                deleteBtnActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(editBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(105, 105, 105)
                .addComponent(updateBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(121, 121, 121)
                .addComponent(deleteBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 161, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(423, 423, 423))
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(38, 38, 38)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel32, javax.swing.GroupLayout.PREFERRED_SIZE, 98, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(jLabel22, javax.swing.GroupLayout.PREFERRED_SIZE, 156, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(manyataRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, 267, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel23, javax.swing.GroupLayout.PREFERRED_SIZE, 156, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel24, javax.swing.GroupLayout.PREFERRED_SIZE, 143, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGap(18, 18, 18)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(tantrikManyatRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, 267, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(tantrikDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 267, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(jLabel21)
                                .addGap(18, 18, 18)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(talukaComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(prashashkiyaDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 267, javax.swing.GroupLayout.PREFERRED_SIZE))))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel26, javax.swing.GroupLayout.PREFERRED_SIZE, 143, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel27, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel34)
                            .addComponent(jLabel39, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel31, javax.swing.GroupLayout.PREFERRED_SIZE, 153, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel30, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel25, javax.swing.GroupLayout.PREFERRED_SIZE, 143, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel29, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                        .addComponent(jLabel18, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel19, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                    .addComponent(jLabel37, javax.swing.GroupLayout.PREFERRED_SIZE, 98, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel33, javax.swing.GroupLayout.PREFERRED_SIZE, 98, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel3Layout.createSequentialGroup()
                                        .addGap(76, 76, 76)
                                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(vibhaagComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(yearComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(kamacheNaavComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(updateyojnaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 267, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                                        .addGap(289, 289, 289)
                                        .addComponent(jLabel35, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel3Layout.createSequentialGroup()
                                        .addGap(45, 45, 45)
                                        .addComponent(sheraComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 339, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(jPanel3Layout.createSequentialGroup()
                                        .addGap(238, 238, 238)
                                        .addComponent(jLabel28, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE)))))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addGap(56, 56, 56)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(kharchWthGstUpd, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(workCompleteDate1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(praptSunUpd, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(kamachiMudatMahineDate1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(akunTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(gstTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(nivedaRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(kamacheAdeshDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(maktedaracheNaavTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(282, 282, 282))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(63, 63, 63)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel19)
                    .addComponent(updateyojnaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel25)
                    .addComponent(maktedaracheNaavTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel18)
                            .addComponent(kamacheNaavComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel26))
                        .addGap(22, 22, 22)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel32)
                            .addComponent(yearComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel27)
                            .addComponent(nivedaRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel33)
                            .addComponent(vibhaagComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel28)))
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addComponent(kamacheAdeshDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(gstTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(18, 18, 18)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(jLabel37)
                                .addGap(27, 27, 27)
                                .addComponent(jLabel21))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(talukaComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(prashashkiyaDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel3Layout.createSequentialGroup()
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel29)
                                    .addComponent(akunTxt1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(kamachiMudatMahineDate1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel30))))
                        .addGap(6, 6, 6)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jLabel31)
                                .addComponent(manyataRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(praptSunUpd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addComponent(jLabel22, javax.swing.GroupLayout.Alignment.TRAILING))
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(16, 16, 16)
                                .addComponent(jLabel34))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addComponent(kharchWthGstUpd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(workCompleteDate1, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(248, 248, 248))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                        .addGap(25, 25, 25)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel23)
                            .addComponent(tantrikDinank1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(tantrikManyatRakkam1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel24)
                            .addComponent(jLabel39))
                        .addGap(31, 31, 31)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(sheraComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel35))
                        .addGap(31, 31, 31)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(editBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(updateBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(deleteBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(99, 99, 99))))
        );

        jTabbedPane1.addTab("माहिती अपडेट करा ", jPanel3);

        jPanel4.setBackground(new java.awt.Color(204, 255, 204));
        jPanel4.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        yojnaNorthTable.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        yojnaNorthTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        yojnaNorthTable.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jScrollPane1.setViewportView(yojnaNorthTable);

        displayYojnaComboBox.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        displayYojnaComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "योजना निवडा" }));

        jLabel20.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        jLabel20.setText("योजना निवडा :");

        showDataBtn.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        showDataBtn.setText("माहिती प्रदर्शित करा");
        showDataBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                showDataBtnActionPerformed(evt);
            }
        });

        jPanel5.setBackground(new java.awt.Color(204, 255, 204));
        jPanel5.setBorder(javax.swing.BorderFactory.createTitledBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true), "कृती", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Mangal", 1, 14))); // NOI18N

        exportToExcel.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        exportToExcel.setText("उत्तर विभाग एक्सेल मध्ये ");
        exportToExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportToExcelActionPerformed(evt);
            }
        });

        closeBtn.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        closeBtn.setText("बंद करा");
        closeBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                closeBtnActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGap(87, 87, 87)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(exportToExcel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(closeBtn, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(71, Short.MAX_VALUE))
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addComponent(exportToExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(closeBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(26, Short.MAX_VALUE))
        );

        abstractTable1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        abstractTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                { new Integer(1), "2", "3", "4", "5", "6", "7", "8", "9"},
                { new Integer(1), "सातारा", null, null, null, null, null, null, null},
                { new Integer(2), "कोरेगाव", null, null, null, null, null, null, null},
                { new Integer(3), "फलटण", null, null, null, null, null, null, null},
                { new Integer(4), "खंडाळा", null, null, null, null, null, null, null},
                { new Integer(5), "वाई", null, null, null, null, null, null, null},
                { new Integer(6), "महाबळेश्वर", null, null, null, null, null, null, null}
            },
            new String [] {
                "अ_क्र", "तालुका", "मंजुर कामे", "निवेदित स्तिथित असले कामांची संख्या", "कार्यारंभ आदेश दिलेल्या कामांची संख्या", "प्रगतीत असलेल्या कामांची संख्या", "पूर्ण झालेल्या कामांची संख्या", "रद्द कामांची संख्या", "शेरा"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        abstractTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jScrollPane2.setViewportView(abstractTable1);

        jPanel6.setBackground(new java.awt.Color(204, 255, 204));
        jPanel6.setBorder(javax.swing.BorderFactory.createTitledBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true), "कृती", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Mangal", 1, 14))); // NOI18N

        exportToExcel1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        exportToExcel1.setText("उत्तर विभाग एक्सेल मध्ये ");
        exportToExcel1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportToExcel1ActionPerformed(evt);
            }
        });

        closeBtn1.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        closeBtn1.setText("बंद करा");
        closeBtn1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                closeBtn1ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel6Layout.createSequentialGroup()
                .addContainerGap(60, Short.MAX_VALUE)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(exportToExcel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(closeBtn1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(83, 83, 83))
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addGap(15, 15, 15)
                .addComponent(exportToExcel1, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(closeBtn1, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        abstractTable2.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        abstractTable2.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                { new Integer(1), "2", "3", "4", "5", "6", "7", "8", "9"},
                { new Integer(1), "माण", null, null, null, null, null, null, null},
                { new Integer(2), "खटाव", null, null, null, null, null, null, null},
                { new Integer(3), "कराड", null, null, null, null, null, null, null},
                { new Integer(4), "पाटण", null, null, null, null, null, null, null},
                { new Integer(5), "जावली", null, null, null, null, null, null, null}
            },
            new String [] {
                "अ_क्र", "तालुका", "मंजुर कामे", "निवेदित स्तिथित असले कामांची संख्या", "कार्यारंभ आदेश दिलेल्या कामांची संख्या", "प्रगतीत असलेल्या कामांची संख्या", "पूर्ण झालेल्या कामांची संख्या", "रद्द कामांची संख्या", "शेरा"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        abstractTable2.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jScrollPane4.setViewportView(abstractTable2);

        jPanel7.setBackground(new java.awt.Color(204, 255, 204));
        jPanel7.setBorder(javax.swing.BorderFactory.createTitledBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true), "कृती", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Mangal", 1, 14))); // NOI18N

        exportToExcel2.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        exportToExcel2.setText("दक्षिण विभाग एक्सेल मध्ये");
        exportToExcel2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportToExcel2ActionPerformed(evt);
            }
        });

        closeBtn2.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        closeBtn2.setText("बंद करा");
        closeBtn2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                closeBtn2ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel7Layout = new javax.swing.GroupLayout(jPanel7);
        jPanel7.setLayout(jPanel7Layout);
        jPanel7Layout.setHorizontalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel7Layout.createSequentialGroup()
                .addGap(69, 69, 69)
                .addGroup(jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(exportToExcel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(closeBtn2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(71, Short.MAX_VALUE))
        );
        jPanel7Layout.setVerticalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel7Layout.createSequentialGroup()
                .addGap(15, 15, 15)
                .addComponent(exportToExcel2, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(closeBtn2, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(21, Short.MAX_VALUE))
        );

        yojnaSouthTable.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        yojnaSouthTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        yojnaSouthTable.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jScrollPane3.setViewportView(yojnaSouthTable);

        jPanel8.setBackground(new java.awt.Color(204, 255, 204));
        jPanel8.setBorder(javax.swing.BorderFactory.createTitledBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true), "कृती", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Mangal", 1, 14))); // NOI18N

        exportToExcel3.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        exportToExcel3.setText("दक्षिण विभाग एक्सेल मध्ये");
        exportToExcel3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportToExcel3ActionPerformed(evt);
            }
        });

        closeBtn3.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        closeBtn3.setText("बंद करा");
        closeBtn3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                closeBtn3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel8Layout = new javax.swing.GroupLayout(jPanel8);
        jPanel8.setLayout(jPanel8Layout);
        jPanel8Layout.setHorizontalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel8Layout.createSequentialGroup()
                .addGap(68, 68, 68)
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(exportToExcel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(closeBtn3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(69, Short.MAX_VALUE))
        );
        jPanel8Layout.setVerticalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel8Layout.createSequentialGroup()
                .addGap(16, 16, 16)
                .addComponent(exportToExcel3, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(closeBtn3, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        yearChooseDisplay.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        yearChooseDisplay.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2021-2022", "2022-2023", "2023-2024", "2024-2025", "2025-2026" }));

        allDataButton.setFont(new java.awt.Font("Mangal", 0, 13)); // NOI18N
        allDataButton.setText("उत्तर आणि दक्षिण विभाग नुसार");
        allDataButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                allDataButtonActionPerformed(evt);
            }
        });

        showMaktedarForm.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N
        showMaktedarForm.setText("मक्तेदार नुसार");
        showMaktedarForm.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                showMaktedarFormActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(14, 14, 14)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addComponent(jLabel20, javax.swing.GroupLayout.PREFERRED_SIZE, 95, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(displayYojnaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 451, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(yearChooseDisplay, javax.swing.GroupLayout.PREFERRED_SIZE, 150, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(showDataBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 155, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(allDataButton)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(showMaktedarForm, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                            .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                            .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jPanel8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addGap(24, 24, 24)
                                .addComponent(jPanel7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)))))
                .addGap(0, 36, Short.MAX_VALUE))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(displayYojnaComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(showDataBtn)
                    .addComponent(yearChooseDisplay, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(allDataButton, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel20)
                    .addComponent(showMaktedarForm, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 277, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 277, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, 277, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 277, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(28, 28, 28)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                        .addComponent(jPanel6, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jPanel8, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jPanel7, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                .addContainerGap(148, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("माहिती प्रदर्शित करा", jPanel4);

        jPanel9.setBackground(new java.awt.Color(204, 255, 204));
        jPanel9.setFont(new java.awt.Font("Mangal", 0, 12)); // NOI18N

        backupBtn.setFont(new java.awt.Font("Mangal", 1, 14)); // NOI18N
        backupBtn.setText("बॅकअप साठी क्लिक करा");
        backupBtn.setBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true));
        backupBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                backupBtnActionPerformed(evt);
            }
        });

        jButton3.setFont(new java.awt.Font("Mangal", 1, 14)); // NOI18N
        jButton3.setText("बंद करा");
        jButton3.setBorder(new javax.swing.border.LineBorder(new java.awt.Color(0, 0, 0), 1, true));
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel9Layout = new javax.swing.GroupLayout(jPanel9);
        jPanel9.setLayout(jPanel9Layout);
        jPanel9Layout.setHorizontalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel9Layout.createSequentialGroup()
                .addGap(367, 367, 367)
                .addComponent(backupBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 234, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(90, 90, 90)
                .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 234, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(481, Short.MAX_VALUE))
        );
        jPanel9Layout.setVerticalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel9Layout.createSequentialGroup()
                .addGap(266, 266, 266)
                .addGroup(jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(backupBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 88, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 88, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(353, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("बॅकअप घ्या", jPanel9);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 1406, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 742, Short.MAX_VALUE)
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void updateyojnaComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_updateyojnaComboBoxActionPerformed
        // TODO add your handling code here:
        try {
            ArrayList<String> works;
            Statement worksStmt;
            ResultSet rsForWorks;
            int i = 0;
            if (updateyojnaComboBox.getSelectedItem().toString().equals("ग्रामीण_रस्ताचा_विकास_वा_मजबुती_करण")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.ग्रामीण_रस्ताचा_विकास_वा_मजबुती_करण";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("इतर_जिल्हा_रस्ते_विकास_व_मजबुतीकरण")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.इतर_जिल्हा_रस्ते_विकास_व_मजबुतीकरण";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("पुर_हानी_व_अतिवष्ठी")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.पुर_हानी_व_अतिवष्ठी";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("शासकीय_जमिनीवरील_अतिक्रमण_रोखण्यासाठी_संरक्षक_भिंत_बांधणे")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.शासकीय_जमिनीवरील_अतिक्रमण_रोखण्यासाठी_संरक्षक_भिंत_बांधणे";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("नावीन्यपूर्ण_योजना")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.नावीन्यपूर्ण_योजना";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("अपारंपारिक_उर्जा")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.अपारंपारिक_उर्जा";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("जिल्हा_नियोजन_मागणी_ओ_283451_सचिवालय_आर्थिक_सेवा")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.जिल्हा_नियोजन_मागणी_ओ_283451_सचिवालय_आर्थिक_सेवा";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("पर्यटन_स्थळ_विकास_मूलभूत_सुविधा_करणे")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.पर्यटन_स्थळ_विकास_मूलभूत_सुविधा_करणे";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));
            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("योजनांचे_मूल्यमापन_सनियंत्रण_व_डाटा_एंट्री_करणे")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.योजनांचे_मूल्यमापन_सनियंत्रण_व_डाटा_एंट्री_करणे";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));

            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("शासकीय_कार्यालयीन_इमारत_बांधकाम")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.शासकीय_कार्यालयीन_इमारत_बांधकाम";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));

            }

            if (updateyojnaComboBox.getSelectedItem().toString().equals("लाँच_खरेदीसाठी_जि_प_ला_अनुदान")) {

                kamacheNaavComboBox.removeAllItems();
                kamacheNaavComboBox.setSelectedItem(null);

                String kaamQuery = "SELECT * FROM एमपीआर.लाँच_खरेदीसाठी_जि_प_ला_अनुदान";
                worksStmt = config.conn.createStatement();
                rsForWorks = worksStmt.executeQuery(kaamQuery);
                works = new ArrayList<>();
                while (rsForWorks.next()) {
                    works.add(rsForWorks.getString("कामाचे_नांव"));
                }

                String[] worksArray = works.toArray(new String[0]);
                kamacheNaavComboBox.setModel(new DefaultComboBoxModel(worksArray));

            }

            selectedTableName = "एमपीआर." + updateyojnaComboBox.getSelectedItem().toString();
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(null, "Database exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_updateyojnaComboBoxActionPerformed

    private void showDataBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_showDataBtnActionPerformed
        // TODO add your handling code here:

        try {

            String tableNameDisplayForTable = displayYojnaComboBox.getSelectedItem().toString();
            tableNameDisplayForTable = "एमपीआर." + tableNameDisplayForTable;
            String yearDisplayForTable = yearChooseDisplay.getSelectedItem().toString();
            String uttarVibhagtalukaArray[] = {"सातारा", "कोरेगाव", "फलटण", "खंडाळा", "वाई", "महाबळेश्वर"};

//            MANJUR KAME UTTAR VIBHAG
            Statement stmtUttarVibhagArray[] = new Statement[uttarVibhagtalukaArray.length];
            ResultSet rsUttarVibhagArray[] = new ResultSet[uttarVibhagtalukaArray.length];
            String uttarVibhagTalukaCount[] = new String[uttarVibhagtalukaArray.length];

            for (int i = 0; i < uttarVibhagtalukaArray.length; i++) {
                stmtUttarVibhagArray[i] = config.conn.createStatement();
                rsUttarVibhagArray[i] = stmtUttarVibhagArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + uttarVibhagtalukaArray[i] + "' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsUttarVibhagArray[i].next()) {
                    uttarVibhagTalukaCount[i] = rsUttarVibhagArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 6; i++) {
                System.out.println(uttarVibhagTalukaCount[i - 1]);
                abstractTable1.setValueAt(uttarVibhagTalukaCount[i - 1], i, 2);
            }

            //NIVIDET STITI UTTAR VIBHAG
            Statement stmtUttarVibhagNividetArray[] = new Statement[uttarVibhagtalukaArray.length];
            ResultSet rsUttarVibhagNividetArray[] = new ResultSet[uttarVibhagtalukaArray.length];
            String uttarVibhagNividetTalukaCount[] = new String[uttarVibhagtalukaArray.length];

            for (int i = 0; i < uttarVibhagtalukaArray.length; i++) {
                stmtUttarVibhagNividetArray[i] = config.conn.createStatement();
                rsUttarVibhagNividetArray[i] = stmtUttarVibhagNividetArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + uttarVibhagtalukaArray[i] + "' AND कामाचा_आदेश_व_दिनांक = N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsUttarVibhagNividetArray[i].next()) {
                    uttarVibhagNividetTalukaCount[i] = rsUttarVibhagNividetArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 6; i++) {
                System.out.println(uttarVibhagNividetTalukaCount[i - 1]);
                abstractTable1.setValueAt(uttarVibhagNividetTalukaCount[i - 1], i, 3);
            }

//            SHERE START UTTAR VIBHAG
            int sheraItemCount = sheraComboBox.getItemCount();
            System.out.println(sheraItemCount);

            String sheraValues[] = new String[sheraItemCount];

            for (int j = 0; j < sheraItemCount; j++) {
                sheraValues[j] = sheraComboBox.getItemAt(j);
            }

            // Print the values to verify
            for (String SHERAVALUES : sheraValues) {
                System.out.println(SHERAVALUES);
            }

            HashMap<String, String> map = new HashMap<>();

            ArrayList mapSataraValues = new ArrayList();
            ArrayList mapKoregaonValues = new ArrayList();
            ArrayList mapPhaltanValues = new ArrayList();
            ArrayList mapKhandalaValues = new ArrayList();
            ArrayList mapWaiValues = new ArrayList();
            ArrayList mapMahabaleshwarValues = new ArrayList();

            Statement sataraShereStmt = config.conn.createStatement();
            ResultSet rsSataraShere;
            String queryShereSatara;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereSatara = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'सातारा' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsSataraShere = sataraShereStmt.executeQuery(queryShereSatara);
                while (rsSataraShere.next()) {
                    if (Integer.parseInt(rsSataraShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsSataraShere.getString("COUNT1") + "-" + sheraValues[i], "सातारा");
                        mapSataraValues.add(rsSataraShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement koregaonShereStmt = config.conn.createStatement();
            ResultSet rsKoregaonShere;
            String queryShereKoregaon;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereKoregaon = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'कोरेगाव' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsKoregaonShere = koregaonShereStmt.executeQuery(queryShereKoregaon);
                while (rsKoregaonShere.next()) {
                    if (Integer.parseInt(rsKoregaonShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsKoregaonShere.getString("COUNT1") + "-" + sheraValues[i], "कोरेगाव");
                        mapKoregaonValues.add(rsKoregaonShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement phaltanShereStmt = config.conn.createStatement();
            ResultSet rsphaltanShere;
            String querySherephaltan;
            for (int i = 0; i < sheraItemCount; i++) {
                querySherephaltan = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'फलटण' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsphaltanShere = phaltanShereStmt.executeQuery(querySherephaltan);
                while (rsphaltanShere.next()) {
                    if (Integer.parseInt(rsphaltanShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsphaltanShere.getString("COUNT1") + "-" + sheraValues[i], "फलटण");
                        mapPhaltanValues.add(rsphaltanShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement khandalaShereStmt = config.conn.createStatement();
            ResultSet rsKhandalaShere;
            String queryShereKhandala;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereKhandala = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'खंडाळा' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsKhandalaShere = khandalaShereStmt.executeQuery(queryShereKhandala);
                while (rsKhandalaShere.next()) {
                    if (Integer.parseInt(rsKhandalaShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsKhandalaShere.getString("COUNT1") + "-" + sheraValues[i], "खंडाळा");
                        mapKhandalaValues.add(rsKhandalaShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement waiShereStmt = config.conn.createStatement();
            ResultSet rsWaiShere;
            String queryShereWai;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereWai = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'वाई' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsWaiShere = waiShereStmt.executeQuery(queryShereWai);
                while (rsWaiShere.next()) {
                    if (Integer.parseInt(rsWaiShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsWaiShere.getString("COUNT1") + "-" + sheraValues[i], "वाई");
                        mapWaiValues.add(rsWaiShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement mahabaleshwarShereStmt = config.conn.createStatement();
            ResultSet rsMahabaleshwarShere;
            String queryShereMahabaleshwar;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereMahabaleshwar = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'महाबळेश्वर' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsMahabaleshwarShere = mahabaleshwarShereStmt.executeQuery(queryShereMahabaleshwar);
                while (rsMahabaleshwarShere.next()) {
                    if (Integer.parseInt(rsMahabaleshwarShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsMahabaleshwarShere.getString("COUNT1") + "-" + sheraValues[i], "महाबळेश्वर");
                        mapMahabaleshwarValues.add(rsMahabaleshwarShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            System.out.println("HashMap Values");
            System.out.println(map);

            System.out.println(mapSataraValues);
            System.out.println(mapKoregaonValues);
            System.out.println(mapPhaltanValues);
            System.out.println(mapKhandalaValues);
            System.out.println(mapWaiValues);
            System.out.println(mapMahabaleshwarValues);

            abstractTable1.setValueAt(mapSataraValues, 1, 8);
            abstractTable1.setValueAt(mapKoregaonValues, 2, 8);
            abstractTable1.setValueAt(mapPhaltanValues, 3, 8);
            abstractTable1.setValueAt(mapKhandalaValues, 4, 8);
            abstractTable1.setValueAt(mapWaiValues, 5, 8);
            abstractTable1.setValueAt(mapMahabaleshwarValues, 6, 8);

//            KARYARAMBH AADESH AND  PRAGATIT KAMANCHI SANKHYA UTTAR VIBHAG
            Statement stmtUttarVibhagKaryaArray[] = new Statement[uttarVibhagtalukaArray.length];
            ResultSet rsUttarVibhagKaryaArray[] = new ResultSet[uttarVibhagtalukaArray.length];
            String uttarVibhagKaryaTalukaCount[] = new String[uttarVibhagtalukaArray.length];

            for (int i = 0; i < uttarVibhagtalukaArray.length; i++) {
                stmtUttarVibhagKaryaArray[i] = config.conn.createStatement();
                rsUttarVibhagKaryaArray[i] = stmtUttarVibhagKaryaArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + uttarVibhagtalukaArray[i] + "' AND कामाचा_आदेश_व_दिनांक != N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsUttarVibhagKaryaArray[i].next()) {
                    uttarVibhagKaryaTalukaCount[i] = rsUttarVibhagKaryaArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 6; i++) {
                System.out.println(uttarVibhagKaryaTalukaCount[i - 1]);
                abstractTable1.setValueAt(uttarVibhagKaryaTalukaCount[i - 1], i, 4);
                abstractTable1.setValueAt(uttarVibhagKaryaTalukaCount[i - 1], i, 5);
            }

            //            PURNA ZALELYA KAMANCHI SANKHYA UTTAR VIBHAG
            Statement stmtUttarVibhagPurnaKaamArray[] = new Statement[uttarVibhagtalukaArray.length];
            ResultSet rsUttarVibhagPurnaKaamArray[] = new ResultSet[uttarVibhagtalukaArray.length];
            String uttarVibhagPurnaKaamTalukaCount[] = new String[uttarVibhagtalukaArray.length];

            for (int i = 0; i < uttarVibhagtalukaArray.length; i++) {
                stmtUttarVibhagPurnaKaamArray[i] = config.conn.createStatement();
                rsUttarVibhagPurnaKaamArray[i] = stmtUttarVibhagPurnaKaamArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + uttarVibhagtalukaArray[i] + "' AND काम_पूर्ण_झालेल्याचे_तारिक != N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsUttarVibhagPurnaKaamArray[i].next()) {
                    uttarVibhagPurnaKaamTalukaCount[i] = rsUttarVibhagPurnaKaamArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 6; i++) {
                System.out.println(uttarVibhagPurnaKaamTalukaCount[i - 1]);
                abstractTable1.setValueAt(uttarVibhagPurnaKaamTalukaCount[i - 1], i, 6);

            }
            for (int i = 1; i <= 6; i++) {
                int a = Integer.parseInt(abstractTable1.getValueAt(i, 5).toString());
                int b = Integer.parseInt(abstractTable1.getValueAt(i, 6).toString());
                int c = a - b;
                abstractTable1.setValueAt(c, i, 5);
            }

            //           RADHHA KAMANCHI SANKHYA UTTAR VIBHAG
            Statement stmtUttarVibhagRadhhaArray[] = new Statement[uttarVibhagtalukaArray.length];
            ResultSet rsUttarVibhagRadhhaArray[] = new ResultSet[uttarVibhagtalukaArray.length];
            String uttarVibhagRadhhaTalukaCount[] = new String[uttarVibhagtalukaArray.length];

            for (int i = 0; i < uttarVibhagtalukaArray.length; i++) {
                stmtUttarVibhagRadhhaArray[i] = config.conn.createStatement();
                rsUttarVibhagRadhhaArray[i] = stmtUttarVibhagRadhhaArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + uttarVibhagtalukaArray[i] + "' AND शेरा = N'काम रद्द' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsUttarVibhagRadhhaArray[i].next()) {
                    uttarVibhagRadhhaTalukaCount[i] = rsUttarVibhagRadhhaArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 6; i++) {
                System.out.println(uttarVibhagRadhhaTalukaCount[i - 1]);
                abstractTable1.setValueAt(uttarVibhagRadhhaTalukaCount[i - 1], i, 7);

            }

            for (int i = 1; i <= 6; i++) {
                int x = Integer.parseInt(abstractTable1.getValueAt(i, 5).toString());
                int y = Integer.parseInt(abstractTable1.getValueAt(i, 7).toString());
                int z = x - y;
                abstractTable1.setValueAt(z, i, 5);
            }

//            DAKSHIN VIBHAAG
            String dakshinVibhagtalukaArray[] = {"माण", "खटाव", "कराड", "पाटण", "जावली"};

            //            MANJUR KAME DAKSHIN VIBHAG
            Statement stmtDakshinVibhagArray[] = new Statement[dakshinVibhagtalukaArray.length];
            ResultSet rsDakshinVibhagArray[] = new ResultSet[dakshinVibhagtalukaArray.length];
            String dakshinVibhagTalukaCount[] = new String[dakshinVibhagtalukaArray.length];

            for (int i = 0; i < dakshinVibhagtalukaArray.length; i++) {
                stmtDakshinVibhagArray[i] = config.conn.createStatement();
                rsDakshinVibhagArray[i] = stmtDakshinVibhagArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + dakshinVibhagtalukaArray[i] + "' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsDakshinVibhagArray[i].next()) {
                    dakshinVibhagTalukaCount[i] = rsDakshinVibhagArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 5; i++) {
                System.out.println(dakshinVibhagTalukaCount[i - 1]);
                abstractTable2.setValueAt(dakshinVibhagTalukaCount[i - 1], i, 2);
            }

            //NIVIDET STITI UTTAR VIBHAG
            Statement stmtDakshinVibhagNividetArray[] = new Statement[dakshinVibhagtalukaArray.length];
            ResultSet rsDakshinVibhagNividetArray[] = new ResultSet[dakshinVibhagtalukaArray.length];
            String dakshinVibhagNividetTalukaCount[] = new String[dakshinVibhagtalukaArray.length];

            for (int i = 0; i < dakshinVibhagtalukaArray.length; i++) {
                stmtDakshinVibhagNividetArray[i] = config.conn.createStatement();
                rsDakshinVibhagNividetArray[i] = stmtDakshinVibhagNividetArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + dakshinVibhagtalukaArray[i] + "' AND कामाचा_आदेश_व_दिनांक = N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsDakshinVibhagNividetArray[i].next()) {
                    dakshinVibhagNividetTalukaCount[i] = rsDakshinVibhagNividetArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 5; i++) {
                System.out.println(dakshinVibhagNividetTalukaCount[i - 1]);
                abstractTable2.setValueAt(dakshinVibhagNividetTalukaCount[i - 1], i, 3);
            }

            //            KARYARAMBH AADESH AND  PRAGATIT KAMANCHI SANKHYA DAKSHIN VIBHAG
            Statement stmtDakshinVibhagKaryaArray[] = new Statement[dakshinVibhagtalukaArray.length];
            ResultSet rsDakshinVibhagKaryaArray[] = new ResultSet[dakshinVibhagtalukaArray.length];
            String dakshinVibhagKaryaTalukaCount[] = new String[dakshinVibhagtalukaArray.length];

            for (int i = 0; i < dakshinVibhagtalukaArray.length; i++) {
                stmtDakshinVibhagKaryaArray[i] = config.conn.createStatement();
                rsDakshinVibhagKaryaArray[i] = stmtDakshinVibhagKaryaArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + dakshinVibhagtalukaArray[i] + "' AND कामाचा_आदेश_व_दिनांक != N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsDakshinVibhagKaryaArray[i].next()) {
                    dakshinVibhagKaryaTalukaCount[i] = rsDakshinVibhagKaryaArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 5; i++) {
                System.out.println(dakshinVibhagKaryaTalukaCount[i - 1]);
                abstractTable2.setValueAt(dakshinVibhagKaryaTalukaCount[i - 1], i, 4);
                abstractTable2.setValueAt(dakshinVibhagKaryaTalukaCount[i - 1], i, 5);  //next4
            }

            //            PURNA ZALELYA KAMANCHI SANKHYA DAKSHIN VIBHAG
            Statement stmtDakshinVibhagPurnaKaamArray[] = new Statement[dakshinVibhagtalukaArray.length];
            ResultSet rsDakshinVibhagPurnaKaamArray[] = new ResultSet[dakshinVibhagtalukaArray.length];
            String dakshinVibhagPurnaKaamTalukaCount[] = new String[dakshinVibhagtalukaArray.length];

            for (int i = 0; i < dakshinVibhagtalukaArray.length; i++) {
                stmtDakshinVibhagPurnaKaamArray[i] = config.conn.createStatement();
                rsDakshinVibhagPurnaKaamArray[i] = stmtDakshinVibhagPurnaKaamArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + dakshinVibhagtalukaArray[i] + "' AND काम_पूर्ण_झालेल्याचे_तारिक != N'NA' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsDakshinVibhagPurnaKaamArray[i].next()) {
                    dakshinVibhagPurnaKaamTalukaCount[i] = rsDakshinVibhagPurnaKaamArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 5; i++) {
                System.out.println(dakshinVibhagPurnaKaamTalukaCount[i - 1]);
                abstractTable2.setValueAt(dakshinVibhagPurnaKaamTalukaCount[i - 1], i, 6);

            }

            for (int i = 1; i <= 5; i++) {
                int a = Integer.parseInt(abstractTable2.getValueAt(i, 5).toString());
                int b = Integer.parseInt(abstractTable2.getValueAt(i, 6).toString());
                int c = a - b;
                abstractTable2.setValueAt(c, i, 5);
            }

            //           RADHHA KAMANCHI SANKHYA DAKSHIN VIBHAG
            Statement stmtDakshinVibhagRadhhaArray[] = new Statement[dakshinVibhagtalukaArray.length];
            ResultSet rsDakshinVibhagRadhhaArray[] = new ResultSet[dakshinVibhagtalukaArray.length];
            String dakshinVibhagRadhhaTalukaCount[] = new String[dakshinVibhagtalukaArray.length];

            for (int i = 0; i < dakshinVibhagtalukaArray.length; i++) {
                stmtDakshinVibhagRadhhaArray[i] = config.conn.createStatement();
                rsDakshinVibhagRadhhaArray[i] = stmtDakshinVibhagRadhhaArray[i].executeQuery("SELECT COUNT(तालुका) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE तालुका=N'" + dakshinVibhagtalukaArray[i] + "' AND शेरा = N'काम रद्द' AND वर्ष=N'" + yearDisplayForTable + "';");
                while (rsDakshinVibhagRadhhaArray[i].next()) {
                    dakshinVibhagRadhhaTalukaCount[i] = rsDakshinVibhagRadhhaArray[i].getString("COUNT1");
                }
            }

            for (int i = 1; i <= 5; i++) {
                System.out.println(dakshinVibhagRadhhaTalukaCount[i - 1]);
                abstractTable2.setValueAt(dakshinVibhagRadhhaTalukaCount[i - 1], i, 7);

            }

            for (int i = 1; i <= 5; i++) {
                int x = Integer.parseInt(abstractTable2.getValueAt(i, 5).toString());
                int y = Integer.parseInt(abstractTable2.getValueAt(i, 7).toString());
                int z = x - y;
                abstractTable2.setValueAt(z, i, 5);
            }

            //SHERE START DAKSHIN VIBHAG
            ArrayList mapMaanValues = new ArrayList();
            ArrayList mapKhatavValues = new ArrayList();
            ArrayList mapKaradValues = new ArrayList();
            ArrayList mapPatanValues = new ArrayList();
            ArrayList mapJaoliValues = new ArrayList();

            Statement maanShereStmt = config.conn.createStatement();
            ResultSet rsMaanShere;
            String queryShereMaan;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereMaan = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'माण' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsMaanShere = maanShereStmt.executeQuery(queryShereMaan);
                while (rsMaanShere.next()) {
                    if (Integer.parseInt(rsMaanShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsMaanShere.getString("COUNT1") + "-" + sheraValues[i], "माण");
                        mapMaanValues.add(rsMaanShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement khatavShereStmt = config.conn.createStatement();
            ResultSet rsKhatavShere;
            String queryShereKhatav;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereKhatav = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'खटाव' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsKhatavShere = khatavShereStmt.executeQuery(queryShereKhatav);
                while (rsKhatavShere.next()) {
                    if (Integer.parseInt(rsKhatavShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsKhatavShere.getString("COUNT1") + "-" + sheraValues[i], "खटाव");
                        mapKhatavValues.add(rsKhatavShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement karadShereStmt = config.conn.createStatement();
            ResultSet rsKaradShere;
            String queryShereKarad;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereKarad = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'कराड' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsKaradShere = karadShereStmt.executeQuery(queryShereKarad);
                while (rsKaradShere.next()) {
                    if (Integer.parseInt(rsKaradShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsKaradShere.getString("COUNT1") + "-" + sheraValues[i], "कराड");
                        mapKaradValues.add(rsKaradShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement patanShereStmt = config.conn.createStatement();
            ResultSet rsPatanShere;
            String querySherePatan;
            for (int i = 0; i < sheraItemCount; i++) {
                querySherePatan = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'पाटण' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsPatanShere = patanShereStmt.executeQuery(querySherePatan);
                while (rsPatanShere.next()) {
                    if (Integer.parseInt(rsPatanShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsPatanShere.getString("COUNT1") + "-" + sheraValues[i], "पाटण");
                        mapPatanValues.add(rsPatanShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            Statement jaoliShereStmt = config.conn.createStatement();
            ResultSet rsJaoliShere;
            String queryShereJaoli;
            for (int i = 0; i < sheraItemCount; i++) {
                queryShereJaoli = "SELECT COUNT(शेरा) AS COUNT1 FROM " + tableNameDisplayForTable + " WHERE शेरा=N'" + sheraValues[i] + "' AND तालुका=N'जावली' AND वर्ष=N'" + yearDisplayForTable + "';";
                rsJaoliShere = jaoliShereStmt.executeQuery(queryShereJaoli);
                while (rsJaoliShere.next()) {
                    if (Integer.parseInt(rsJaoliShere.getString("COUNT1")) >= 1) {
//                            sheraUttarVibhagCount[i] = uttarVibhagtalukaArray[j] + " - " + sheraValues[i] + " - " + rsUttarVibhagShera[i].getString("COUNT(शेरा)");

                        map.put(rsJaoliShere.getString("COUNT1") + "-" + sheraValues[i], "जावली");
                        mapJaoliValues.add(rsJaoliShere.getString("COUNT1") + "-" + sheraValues[i]);
                    }
                }
            }

            abstractTable2.setValueAt(mapMaanValues, 1, 8);
            abstractTable2.setValueAt(mapKhatavValues, 2, 8);
            abstractTable2.setValueAt(mapKaradValues, 3, 8);
            abstractTable2.setValueAt(mapPatanValues, 4, 8);
            abstractTable2.setValueAt(mapJaoliValues, 5, 8);

            rs = config.conn.getMetaData().getColumns(null, null, displayYojnaComboBox.getSelectedItem().toString(), "%");
            TABLE_COLUMNS = new ArrayList<>();
            TABLE_COLUMNS2 = new ArrayList<>();
            while (rs.next()) {
                String column = rs.getString("COLUMN_NAME");
                TABLE_COLUMNS.add(column);
                TABLE_COLUMNS2.add(column);
            }
            prepTableSpec(yojnaNorthTable, yojnaSouthTable, yearDisplayForTable);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(), "Print Data", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }
    }//GEN-LAST:event_showDataBtnActionPerformed

    private void exportToExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportToExcelActionPerformed
        // TODO add your handling code here:
        exportToExcel(yojnaNorthTable);
    }//GEN-LAST:event_exportToExcelActionPerformed

    private void closeBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_closeBtnActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_closeBtnActionPerformed

    private void kamacheNaavComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_kamacheNaavComboBoxActionPerformed
        try {
            // TODO add your handling code here:
            String workName = kamacheNaavComboBox.getSelectedItem().toString();
//            String table=updateyojnaComboBox.getSelectedItem().toString();
            System.out.println(selectedTableName);
            String querySelect = "SELECT * FROM " + selectedTableName + " WHERE कामाचे_नांव = N'" + workName + "';";
            Statement queryStmt = config.conn.createStatement();
            ResultSet rsForUpdate = queryStmt.executeQuery(querySelect);

            if (rsForUpdate.next()) {
                yearComboBox1.setSelectedItem(rsForUpdate.getString("वर्ष"));
                vibhaagComboBox1.setSelectedItem(rsForUpdate.getString("बांधकाम_विभाग"));
                talukaComboBox1.setSelectedItem(rsForUpdate.getString("तालुका"));
                if ((rsForUpdate.getString("प्रशासकीय_मान्यता_दिनाक")).equals("NA")) {
                    System.out.println("Hello null");
                    prashashkiyaDinank1.setDate(null);
                } else {
                    System.out.println("Hello Not null");
                    prashashkiyaDinank1.setDate(new SimpleDateFormat("yyyy-MM-dd").parse(rsForUpdate.getString("प्रशासकीय_मान्यता_दिनाक")));
                }
                manyataRakkam1.setText(rsForUpdate.getString("प्रशासकीय_मान्यता_रक्कम"));
                if ((rsForUpdate.getString("तांत्रीक_मान्यता_दिनांक")).equals("NA")) {
                    System.out.println("Hello null");
                    tantrikDinank1.setDate(null);
                } else {
                    System.out.println("Hello Not null");
                    tantrikDinank1.setDate(new SimpleDateFormat("yyyy-MM-dd").parse(rsForUpdate.getString("तांत्रीक_मान्यता_दिनांक")));
                }
                tantrikManyatRakkam1.setText(rsForUpdate.getString("तांत्रीक_मान्यता_रक्कम"));
                maktedaracheNaavTxt1.setText(rsForUpdate.getString("मक्तेदाराचे_नांव"));
                if ((rsForUpdate.getString("कामाचा_आदेश_व_दिनांक")).equals("NA")) {
                    System.out.println("Hello null");
                    kamacheAdeshDinank1.setDate(null);
                } else {
                    System.out.println("Hello Not null");
                    kamacheAdeshDinank1.setDate(new SimpleDateFormat("yyyy-MM-dd").parse(rsForUpdate.getString("कामाचा_आदेश_व_दिनांक")));
                }
                nivedaRakkam1.setText(rsForUpdate.getString("निविदा_स्विकृती_रक्कम"));
                gstTxt1.setText(rsForUpdate.getString("जी_एस_टी"));
                akunTxt1.setText(rsForUpdate.getString("एकुण"));
//                kamachiMudatMahineDate1.setText(rsForUpdate.getString("कामाची_मुदत_महिने"));
                if ((rsForUpdate.getString("कामाची_मुदत_महिने")).equals("NA")) {
                    System.out.println("Hello null");
                    kamachiMudatMahineDate1.setDate(null);
                } else {
                    System.out.println("Hello Not null");
                    kamachiMudatMahineDate1.setDate(new SimpleDateFormat("yyyy-MM-dd").parse(rsForUpdate.getString("कामाची_मुदत_महिने")));
                }
                praptSunUpd.setText(rsForUpdate.getString("प्राप्त_निधी_या_वर्षा_मध्ये"));

                kharchWthGstUpd.setText(rsForUpdate.getString("या_वर्षा_मधील_खर्च_जी_एस_टी_सह"));

                if ((rsForUpdate.getString("काम_पूर्ण_झालेल्याचे_तारिक")).equals("NA")) {
                    System.out.println("Hello null");
                    workCompleteDate1.setDate(null);
                } else {
                    System.out.println("Hello Not null");
                    workCompleteDate1.setDate(new SimpleDateFormat("yyyy-MM-dd").parse(rsForUpdate.getString("काम_पूर्ण_झालेल्याचे_तारिक")));
                }
                sheraComboBox1.setSelectedItem(rsForUpdate.getString("शेरा"));

                deleteBtn.setEnabled(true);
            }

        } catch (Exception ex) {
//            JOptionPane.showMessageDialog(null, ex.toString(), "Monthly Progress Report", JOptionPane.WARNING_MESSAGE);
            ex.printStackTrace();
        }

    }//GEN-LAST:event_kamacheNaavComboBoxActionPerformed

    private void updateBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_updateBtnActionPerformed
        try {
            // TODO add your handling code here:

            tableName = "एमपीआर." + updateyojnaComboBox.getSelectedItem().toString();
            kamacheNaavTxtStr = kamacheNaavComboBox.getSelectedItem().toString();
            year = yearComboBox1.getSelectedItem().toString();
            bandkamVibhag = vibhaagComboBox1.getSelectedItem().toString();
            taluka = talukaComboBox1.getSelectedItem().toString();
            if (prashashkiyaDinank1.getDate() != null) {
                prashashkiyaDinankStr = toDataBaseDate.format(prashashkiyaDinank1.getDate());
            } else {
                prashashkiyaDinankStr = "NA";
            }
            if (manyataRakkam1.getText().isBlank()) {
                manyataRakkamStr = "0";
            } else {
                manyataRakkamStr = manyataRakkam1.getText().toString();
            }
            if (tantrikDinank1.getDate() != null) {
                tantrikDinankStr = toDataBaseDate.format(tantrikDinank1.getDate());
            } else {
                tantrikDinankStr = "NA";
            }
            tantrikManyatRakkamStr = tantrikManyatRakkam1.getText().toString();
            maktedaracheNaavTxtStr = maktedaracheNaavTxt1.getText().toString();
            if (kamacheAdeshDinank1.getDate() != null) {
                kamacheAdeshDinankStr = toDataBaseDate.format(kamacheAdeshDinank1.getDate());
            } else {
                kamacheAdeshDinankStr = "NA";
            }
            nivedaRakkamStr = nivedaRakkam1.getText().toString();
            gstTxtStr = gstTxt1.getText().toString();
            akunTxtStr = akunTxt1.getText().toString();
            if (kamachiMudatMahineDate1.getDate() != null) {
                kamachiMudatMahineStr = toDataBaseDate.format(kamachiMudatMahineDate1.getDate());
            } else {
                kamachiMudatMahineStr = "NA";
            }
            praptSunStr = praptSunUpd.getText().toString();

            if (kharchWthGstUpd.getText().isBlank()) {
                kharchWthGstStr = "0";
            } else {
                kharchWthGstStr = kharchWthGstUpd.getText().toString();
            }

            if (workCompleteDate1.getDate() != null) {
                workCompleteDateStr = toDataBaseDate.format(workCompleteDate1.getDate());
            } else {
                workCompleteDateStr = "NA";
            }

            sheraStr = sheraComboBox1.getSelectedItem().toString();

            PreparedStatement preparedStatement = null;

            String sql = "UPDATE " + tableName + " SET वर्ष=?,बांधकाम_विभाग=?,तालुका=?,कामाचे_नांव = ?, प्रशासकीय_मान्यता_दिनाक = ?, प्रशासकीय_मान्यता_रक्कम = ?, "
                    + "तांत्रीक_मान्यता_दिनांक = ?, तांत्रीक_मान्यता_रक्कम = ?, मक्तेदाराचे_नांव = ?, कामाचा_आदेश_व_दिनांक = ?, "
                    + "निविदा_स्विकृती_रक्कम = ?, जी_एस_टी = ?, एकुण = ?, कामाची_मुदत_महिने = ?, "
                    + "प्राप्त_निधी_या_वर्षा_मध्ये = ?, या_वर्षा_मधील_खर्च_जी_एस_टी_सह = ?, काम_पूर्ण_झालेल्याचे_तारिक = ?, शेरा = ? "
                    + "WHERE कामाचे_नांव = ?";

            // Prepare the statement
            preparedStatement = config.conn.prepareStatement(sql);
            preparedStatement.setString(1, year);
            preparedStatement.setString(2, bandkamVibhag);
            preparedStatement.setString(3, taluka);
            preparedStatement.setString(4, kamacheNaavTxtStr);
            preparedStatement.setString(5, prashashkiyaDinankStr);
            preparedStatement.setString(6, manyataRakkamStr);
            preparedStatement.setString(7, tantrikDinankStr);
            preparedStatement.setString(8, tantrikManyatRakkamStr);
            preparedStatement.setString(9, maktedaracheNaavTxtStr);
            preparedStatement.setString(10, kamacheAdeshDinankStr);
            preparedStatement.setString(11, nivedaRakkamStr);
            preparedStatement.setString(12, gstTxtStr);
            preparedStatement.setString(13, akunTxtStr);
            preparedStatement.setString(14, kamachiMudatMahineStr);
            preparedStatement.setString(15, praptSunStr);
            preparedStatement.setString(16, kharchWthGstStr);
            preparedStatement.setString(17, workCompleteDateStr);
            preparedStatement.setString(18, sheraStr);
            preparedStatement.setString(19, kamacheNaavTxtStr); // Set the condition value here

            // Execute the update
            int rowsAffected = preparedStatement.executeUpdate();
            System.out.println("Rows affected: " + rowsAffected);
            if (rowsAffected >= 1) {
                JOptionPane.showMessageDialog(null, "माहिती योग्यरीतीने अपडेट झाली", "Monthly Progress Report", JOptionPane.INFORMATION_MESSAGE);
                disableUpdateFields(false);
                clearAllFields();
                updateBtn.setEnabled(false);
                deleteBtn.setEnabled(false);
            } else {
                JOptionPane.showMessageDialog(null, "माहिती अपडेट झाली नाही", "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }


    }//GEN-LAST:event_updateBtnActionPerformed

    private void editBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_editBtnActionPerformed
        // TODO add your handling code here:
        disableUpdateFields(true);
        updateBtn.setEnabled(true);
    }//GEN-LAST:event_editBtnActionPerformed

    private void saveBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveBtnActionPerformed
        try {
            // TODO add your handling code here:
            year = yearComboBox.getSelectedItem().toString();
            bandkamVibhag = vibhaagComboBox.getSelectedItem().toString();
            taluka = talukaComboBox.getSelectedItem().toString();
            tableName = "एमपीआर." + yojnaComboBox.getSelectedItem().toString();
            kamacheNaavTxtStr = kamacheNaavTxt.getText().toString();

            if (prashashkiyaDinank.getDate() != null) {
                prashashkiyaDinankStr = toDataBaseDate.format(prashashkiyaDinank.getDate());
            } else {
                prashashkiyaDinankStr = "NA";
            }

            if (manyataRakkam.getText().isBlank()) {
                manyataRakkamStr = "0";

            } else {
                manyataRakkamStr = manyataRakkam.getText().toString();
            }

            if (tantrikDinank.getDate() != null) {
                tantrikDinankStr = toDataBaseDate.format(tantrikDinank.getDate());
            } else {
                tantrikDinankStr = "NA";
            }

            tantrikManyatRakkamStr = tantrikManyatRakkam.getText().toString();
            maktedaracheNaavTxtStr = maktedaracheNaavTxt.getText().toString();
            if (kamacheAdeshDinank.getDate() != null) {
                kamacheAdeshDinankStr = toDataBaseDate.format(kamacheAdeshDinank.getDate());
            } else {
                kamacheAdeshDinankStr = "NA";
            }
            nivedaRakkamStr = nivedaRakkam.getText().toString();
            gstTxtStr = gstTxt.getText().toString();
            akunTxtStr = akunTxt.getText().toString();
            if (kamachiMudatMahineDate.getDate() != null) {
                kamachiMudatMahineStr = toDataBaseDate.format(kamachiMudatMahineDate.getDate());
            } else {
                kamachiMudatMahineStr = "NA";
            }
            praptSunStr = praptSun.getText().toString();
            if (kharchWthGst.getText().isBlank()) {
                kharchWthGstStr = "0";

            } else {
                kharchWthGstStr = kharchWthGst.getText().toString();
            }
            if (workCompleteDate.getDate() != null) {
                workCompleteDateStr = toDataBaseDate.format(workCompleteDate.getDate());
            } else {
                workCompleteDateStr = "NA";
            }

            sheraStr = sheraComboBox.getSelectedItem().toString();

            String insertQuery = "INSERT INTO " + tableName + " (वर्ष,बांधकाम_विभाग,तालुका,कामाचे_नांव, प्रशासकीय_मान्यता_दिनाक, प्रशासकीय_मान्यता_रक्कम, तांत्रीक_मान्यता_दिनांक, तांत्रीक_मान्यता_रक्कम, मक्तेदाराचे_नांव, कामाचा_आदेश_व_दिनांक, निविदा_स्विकृती_रक्कम, जी_एस_टी, एकुण, कामाची_मुदत_महिने, प्राप्त_निधी_या_वर्षा_मध्ये, या_वर्षा_मधील_खर्च_जी_एस_टी_सह, काम_पूर्ण_झालेल्याचे_तारिक, शेरा) VALUES(N'" + year + "',N'" + bandkamVibhag + "',N'" + taluka + "',N'" + kamacheNaavTxtStr + "',N'" + prashashkiyaDinankStr + "',N'" + manyataRakkamStr + "',N'" + tantrikDinankStr + "',N'" + tantrikManyatRakkamStr + "',N'" + maktedaracheNaavTxtStr + "', N'" + kamacheAdeshDinankStr + "', N'" + nivedaRakkamStr + "', N'" + gstTxtStr + "', N'" + akunTxtStr + "', N'" + kamachiMudatMahineStr + "', N'" + praptSunStr + "', N'" + kharchWthGstStr + "',N'" + workCompleteDateStr + "', N'" + sheraStr + "');";
            System.out.println(tableName);
            System.out.println(kamacheNaavTxtStr);
            System.out.println(prashashkiyaDinankStr);
            System.out.println(manyataRakkamStr);
            System.out.println(tantrikDinankStr);
            System.out.println(tantrikManyatRakkamStr);
            System.out.println(insertQuery);
            Statement insertStmt = config.conn.createStatement();
            int insertInt = insertStmt.executeUpdate(insertQuery);

            if (insertInt >= 1) {
                JOptionPane.showMessageDialog(null, "माहिती योग्यरीतीने दाखल झाली", "Monthly Progress Report", JOptionPane.INFORMATION_MESSAGE);
                clearAllFields();
            } else {
                JOptionPane.showMessageDialog(null, "माहिती दाखल झाली नाही", "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

            }
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(null, "Database exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

        }

    }//GEN-LAST:event_saveBtnActionPerformed

    private void vibhaagComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_vibhaagComboBoxActionPerformed
        // TODO add your handling code here:
        if (vibhaagComboBox.getSelectedItem().toString().equals("उत्तर विभाग")) {

            talukaComboBox.removeAllItems();
            talukaComboBox.setSelectedItem(null);
            talukaComboBox.addItem("सातारा");
            talukaComboBox.addItem("कोरेगाव");
            talukaComboBox.addItem("फलटण");
            talukaComboBox.addItem("खंडाळा");
            talukaComboBox.addItem("वाई");
            talukaComboBox.addItem("महाबळेश्वर");

        }

        if (vibhaagComboBox.getSelectedItem().toString().equals("दक्षिण विभाग")) {

            talukaComboBox.removeAllItems();
            talukaComboBox.setSelectedItem(null);
            talukaComboBox.addItem("माण");
            talukaComboBox.addItem("खटाव");
            talukaComboBox.addItem("कराड");
            talukaComboBox.addItem("पाटण");
            talukaComboBox.addItem("जावली");

        }
    }//GEN-LAST:event_vibhaagComboBoxActionPerformed

    private void vibhaagComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_vibhaagComboBox1ActionPerformed
        // TODO add your handling code here:

        if (vibhaagComboBox1.getSelectedItem().toString().equals("उत्तर विभाग")) {

            talukaComboBox1.removeAllItems();
            talukaComboBox1.setSelectedItem(null);
            talukaComboBox1.addItem("सातारा");
            talukaComboBox1.addItem("कोरेगाव");
            talukaComboBox1.addItem("फलटण");
            talukaComboBox1.addItem("खंडाळा");
            talukaComboBox1.addItem("वाई");
            talukaComboBox1.addItem("महाबळेश्वर");

        }

        if (vibhaagComboBox1.getSelectedItem().toString().equals("दक्षिण विभाग")) {

            talukaComboBox1.removeAllItems();
            talukaComboBox1.setSelectedItem(null);
            talukaComboBox1.addItem("माण");
            talukaComboBox1.addItem("खटाव");
            talukaComboBox1.addItem("कराड");
            talukaComboBox1.addItem("पाटण");
            talukaComboBox1.addItem("जावली");

        }

    }//GEN-LAST:event_vibhaagComboBox1ActionPerformed

    private void exportToExcel1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportToExcel1ActionPerformed
        // TODO add your handling code here:
        exportToExcelUttar(abstractTable1);
    }//GEN-LAST:event_exportToExcel1ActionPerformed

    private void closeBtn1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_closeBtn1ActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_closeBtn1ActionPerformed

    private void exportToExcel2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportToExcel2ActionPerformed
        // TODO add your handling code here:
        exportToExcelUttar(abstractTable2);
    }//GEN-LAST:event_exportToExcel2ActionPerformed

    private void closeBtn2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_closeBtn2ActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_closeBtn2ActionPerformed

    private void closeBtn3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_closeBtn3ActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_closeBtn3ActionPerformed

    private void exportToExcel3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportToExcel3ActionPerformed
        // TODO add your handling code here:
        exportToExcel(yojnaSouthTable);
    }//GEN-LAST:event_exportToExcel3ActionPerformed

    private void kamachiMudatMahineDateFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_kamachiMudatMahineDateFocusGained
        // TODO add your handling code here:

    }//GEN-LAST:event_kamachiMudatMahineDateFocusGained

    private void kamachiMudatMahineDateMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_kamachiMudatMahineDateMouseClicked
        // TODO add your handling code here:

    }//GEN-LAST:event_kamachiMudatMahineDateMouseClicked

    private void praptSunFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_praptSunFocusGained
        // TODO add your handling code here:
        nivedaRakkamLong = Long.parseLong(nivedaRakkam.getText().toString());
        gstTxtLong = Long.parseLong(gstTxt.getText().toString());
        long totalLong = nivedaRakkamLong + gstTxtLong;
        akunTxt.setText(Long.toString(totalLong));
        System.out.println(totalLong);
    }//GEN-LAST:event_praptSunFocusGained

    private void praptSunUpdFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_praptSunUpdFocusGained
        // TODO add your handling code here:
        nivedaRakkamLong1 = Long.parseLong(nivedaRakkam1.getText().toString());
        gstTxtLong1 = Long.parseLong(gstTxt1.getText().toString());
        long totalLong1 = nivedaRakkamLong1 + gstTxtLong1;
        akunTxt1.setText(Long.toString(totalLong1));
        System.out.println(totalLong1);
    }//GEN-LAST:event_praptSunUpdFocusGained

    private void deleteBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deleteBtnActionPerformed
        try {

            int result = JOptionPane.showConfirmDialog(null, "आपण कामाचे रेकॉर्ड हटवू इच्छित आहात याची खात्री आहे का?", "Pension Record System", JOptionPane.YES_OPTION);
            if (result == 0) {

                String tableNameDelete = "एमपीआर." + updateyojnaComboBox.getSelectedItem().toString();
                String workNameDelete = kamacheNaavComboBox.getSelectedItem().toString();

                Statement deleteStmt = config.conn.createStatement();
                String queryDelete = "DELETE from " + tableNameDelete + " WHERE कामाचे_नांव=N'" + workNameDelete + "';";

                int deleteInt = deleteStmt.executeUpdate(queryDelete);
                if (deleteInt > 0) {
                    JOptionPane.showMessageDialog(null, "माहिती योग्यरीतीने डिलीट झाली आहे", "Monthly Progress Report", JOptionPane.INFORMATION_MESSAGE);
                    disableUpdateFields(false);
                    clearAllFields();
                    updateBtn.setEnabled(false);
                    deleteBtn.setEnabled(false);
                } else {
                    JOptionPane.showMessageDialog(null, "Some error occured", "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

                }
            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

        }


    }//GEN-LAST:event_deleteBtnActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void allDataButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_allDataButtonActionPerformed
        // TODO add your handling code here:
        AllData allData = new AllData();
        allData.setVisible(true);

    }//GEN-LAST:event_allDataButtonActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        // TODO add your handling code here:
        Login login = new Login();
        login.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton3ActionPerformed

    private void backupBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_backupBtnActionPerformed
        // TODO add your handling code here:

        JFileChooser path = new JFileChooser();
        path.showOpenDialog(this);
        Process p = null;
        String date = new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date());
        try {
            File f = path.getSelectedFile();
            location = f.getAbsolutePath();
            System.out.println(location);
            location = location.replace('\\', '/');
            filename = location + "_" + date + ".bak";

            System.out.println(location);

            System.out.println(filename);

            String sql = "BACKUP DATABASE एमपीआर2 TO DISK = '" + filename + "'";
            Statement statement = config.conn.createStatement();
            int flag = statement.executeUpdate(sql);

            System.out.println(flag);
            if (flag >= -1) {
                JOptionPane.showMessageDialog(null, "Backup Created", "Pension Record System", JOptionPane.INFORMATION_MESSAGE);

            } else {
                JOptionPane.showMessageDialog(null, "Backup Failed", "Pension Record System", JOptionPane.ERROR_MESSAGE);
            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(), "Pension Record System", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }

    }//GEN-LAST:event_backupBtnActionPerformed

    private void showMaktedarFormActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_showMaktedarFormActionPerformed
        // TODO add your handling code here:
        MaktedarData maktedar = new MaktedarData();
        maktedar.setVisible(true);
    }//GEN-LAST:event_showMaktedarFormActionPerformed

    private void exportToExcel(JTable jt) {
        try {
            JFileChooser jFileChooser = new JFileChooser();
            jFileChooser.showSaveDialog(jt);
            File saveFile = jFileChooser.getSelectedFile();
            if (saveFile != null) {
                saveFile = new File(saveFile.toString() + ".xlsx");
                Workbook wb = new XSSFWorkbook();
                Sheet sheet = wb.createSheet(displayYojnaComboBox.getSelectedItem().toString() + "_" + yearChooseDisplay.getSelectedItem().toString());
                sheet.setColumnWidth(0, 1500);

                Row rowHeader = sheet.createRow(0);

                Row rowCol = sheet.createRow(1);
                XSSFFont font = ((XSSFWorkbook) wb).createFont();
                font.setBold(true);

                CellStyle style = wb.createCellStyle();
                style.setFont(font);
                style.setWrapText(true);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);

                XSSFFont fontHeader = ((XSSFWorkbook) wb).createFont();
                fontHeader.setBold(true);
                fontHeader.setFontHeightInPoints((short) 20);
                CellStyle styleHeader = wb.createCellStyle();
                styleHeader.setFont(fontHeader);
//                styleHeader.setWrapText(true);

                Cell cellHeader = rowHeader.createCell(6);

                cellHeader.setCellStyle(styleHeader);
                cellHeader.setCellValue(displayYojnaComboBox.getSelectedItem().toString());

                for (int i = 0; i < jt.getColumnCount(); i++) {
                    Cell cell = rowCol.createCell(i);
                    cell.setCellStyle(style);
                    cell.setCellValue(jt.getColumnName(i));
                    sheet.setColumnWidth(i + 1, 4000);

                }
                sheet.setRepeatingRows(CellRangeAddress.valueOf("2:2"));
                Cell cell = null;

                CellStyle styleData = wb.createCellStyle();
                styleData.setWrapText(true);
                styleData.setBorderTop(BorderStyle.THIN);
                styleData.setBorderBottom(BorderStyle.THIN);
                styleData.setBorderLeft(BorderStyle.THIN);
                styleData.setBorderRight(BorderStyle.THIN);

                for (int j = 0; j < jt.getRowCount(); j++) {
                    Row row = sheet.createRow(j + 2);
                    for (int k = 0; k < jt.getColumnCount(); k++) {
                        cell = row.createCell(k);
                        if (jt.getValueAt(j, k) != null) {
                            if (j == 0) {
                                cell.setCellStyle(style);
                            }
                            cell.setCellStyle(styleData);
                            cell.setCellValue(jt.getValueAt(j, k).toString());

                        }

                    }
                }

                Row rowTotal = sheet.createRow(jt.getRowCount() + 2);

                Cell cellTotal = rowTotal.createCell(0);
                cellTotal.setCellValue("एकूण");
                cellTotal.setCellStyle(style);

                long prashashkiyaSum = 0;
                long kharchWithGstSum = 0;
                for (int i = 1; i < jt.getRowCount(); i++) {
                    prashashkiyaSum = (long) (prashashkiyaSum + Long.parseLong(jt.getValueAt(i, 5).toString()));

                    kharchWithGstSum = (long) (kharchWithGstSum + Integer.parseInt(jt.getValueAt(i, 15).toString()));
                }
                System.out.println("prashashkiyaSum Sum : " + prashashkiyaSum);

                Cell prashashkiyaCell = rowTotal.createCell(5);
                prashashkiyaCell.setCellValue(Long.toString(prashashkiyaSum));
                prashashkiyaCell.setCellStyle(style);

                Cell kharchWithGstCell = rowTotal.createCell(15);
                kharchWithGstCell.setCellValue(Long.toString(kharchWithGstSum));
                kharchWithGstCell.setCellStyle(style);

                
                 sheet.setFitToPage(true); 

                PrintSetup printSetup = sheet.getPrintSetup();
                printSetup.setLandscape(true);
//                printSetup.setScale((short) 95); // Set scale to 95%
//               // Enable fit to page
//                printSetup.setFitWidth((short) 1); // Fit all columns to one page
//                printSetup.setFitHeight((short) 0);
                
               printSetup.setPaperSize(PrintSetup.LETTER_PAPERSIZE);
                

                FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                wb.write(out);
                wb.close();
                out.close();
                openFile(saveFile.toString());

            } else {
                JOptionPane.showMessageDialog(null, "Operation Cancelled", "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

        }
    }

    public void exportToExcelUttar(JTable jt) {
        try {
            JFileChooser jFileChooser = new JFileChooser();
            jFileChooser.showSaveDialog(jt);
            File saveFile = jFileChooser.getSelectedFile();
            if (saveFile != null) {
                saveFile = new File(saveFile.toString() + ".xlsx");
                Workbook wb = new XSSFWorkbook();
                Sheet sheet = wb.createSheet(displayYojnaComboBox.getSelectedItem().toString() + "_" + yearChooseDisplay.getSelectedItem().toString());
                sheet.setColumnWidth(0, 1500);
                Row rowCol = sheet.createRow(0);
                XSSFFont font = ((XSSFWorkbook) wb).createFont();
                font.setBold(true);
                CellStyle style = wb.createCellStyle();
                style.setFont(font);
                for (int i = 0; i < jt.getColumnCount(); i++) {
                    Cell cell = rowCol.createCell(i);
                    cell.setCellStyle(style);
                    cell.setCellValue(jt.getColumnName(i));
                    sheet.setColumnWidth(i + 1, 4000);

                }
                Cell cell = null;
                for (int j = 0; j < jt.getRowCount(); j++) {
                    Row row = sheet.createRow(j + 1);
                    for (int k = 0; k < jt.getColumnCount(); k++) {
                        cell = row.createCell(k);
                        if (jt.getValueAt(j, k) != null) {
                            if (j == 0) {
                                cell.setCellStyle(style);
                            }
                            cell.setCellValue(jt.getValueAt(j, k).toString());

                        }

                    }
                }

                Row rowTotal = sheet.createRow(jt.getRowCount() + 1);

                Cell cellTotal = rowTotal.createCell(0);
                cellTotal.setCellValue("एकूण");
                cellTotal.setCellStyle(style);

                int manjurSum = 0;
                int nividitSum = 0;
                int karyarambhAdeshSum = 0;
                int pragatitKaamSum = 0;
                int purnaZaleliKaamSum = 0;
                int raddhaKaamSum = 0;
                for (int i = 1; i < jt.getRowCount(); i++) {
                    manjurSum = (int) (manjurSum + Integer.parseInt(jt.getValueAt(i, 2).toString()));
                    nividitSum = (int) (nividitSum + Integer.parseInt(jt.getValueAt(i, 3).toString()));;
                    karyarambhAdeshSum = (int) (karyarambhAdeshSum + Integer.parseInt(jt.getValueAt(i, 4).toString()));
                    pragatitKaamSum = (int) (pragatitKaamSum + Integer.parseInt(jt.getValueAt(i, 5).toString()));
                    purnaZaleliKaamSum = (int) (purnaZaleliKaamSum + Integer.parseInt(jt.getValueAt(i, 6).toString()));
                    raddhaKaamSum = (int) (raddhaKaamSum + Integer.parseInt(jt.getValueAt(i, 7).toString()));
                }
                System.out.println("Manjur Sum : " + manjurSum);

                Cell manjurCell = rowTotal.createCell(2);
                manjurCell.setCellValue(Integer.toString(manjurSum));
                manjurCell.setCellStyle(style);

                Cell nividitCell = rowTotal.createCell(3);
                nividitCell.setCellValue(Integer.toString(nividitSum));
                nividitCell.setCellStyle(style);

                Cell karyarambhAdeshCell = rowTotal.createCell(4);
                karyarambhAdeshCell.setCellValue(Integer.toString(karyarambhAdeshSum));
                karyarambhAdeshCell.setCellStyle(style);

                Cell pragatitKaamCell = rowTotal.createCell(5);
                pragatitKaamCell.setCellValue(Integer.toString(pragatitKaamSum));
                pragatitKaamCell.setCellStyle(style);

                Cell purnaZaleliKaamCell = rowTotal.createCell(6);
                purnaZaleliKaamCell.setCellValue(Integer.toString(purnaZaleliKaamSum));
                purnaZaleliKaamCell.setCellStyle(style);

                Cell raddhaKaamCell = rowTotal.createCell(7);
                raddhaKaamCell.setCellValue(Integer.toString(raddhaKaamSum));
                raddhaKaamCell.setCellStyle(style);

                FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                wb.write(out);
                wb.close();
                out.close();
                openFile(saveFile.toString());

            } else {
                JOptionPane.showMessageDialog(null, "Operation Cancelled", "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

        }
    }

    public static void openFile(String file) {
        try {
            File path = new File(file);
            Desktop.getDesktop().open(path);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Exception occured : " + ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);

        }
    }

    public void prepTableSpec(JTable dataTable, JTable dataTable2, String year) {
        model = (DefaultTableModel) dataTable.getModel();
        model.getDataVector().removeAllElements();
        model.fireTableStructureChanged();

        model2 = (DefaultTableModel) dataTable2.getModel();
        model2.getDataVector().removeAllElements();
        model2.fireTableStructureChanged();

        if (dataTable.getCellEditor() != null) {
            dataTable.getCellEditor().stopCellEditing();
        }
        String[] columns = TABLE_COLUMNS.stream().toArray(size -> new String[size]);
        model = new DefaultTableModel(columns, 0) {

            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }

        };

        dataTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        dataTable.getSelectionModel().addListSelectionListener((ListSelectionEvent e) -> {

            if (!e.getValueIsAdjusting()) {

            }

        });
        dataTable.setModel(model);

        //NEW
        if (dataTable2.getCellEditor() != null) {
            dataTable2.getCellEditor().stopCellEditing();
        }
        String[] columns2 = TABLE_COLUMNS2.stream().toArray(size -> new String[size]);
        model2 = new DefaultTableModel(columns2, 0) {

            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }

        };

        dataTable2.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        dataTable2.getSelectionModel().addListSelectionListener((ListSelectionEvent e) -> {

            if (!e.getValueIsAdjusting()) {

            }

        });
        dataTable2.setModel(model2);

        //New
        try {

            String SQL = "SELECT * FROM एमपीआर." + displayYojnaComboBox.getSelectedItem().toString() + " WHERE बांधकाम_विभाग = N'उत्तर विभाग' AND वर्ष=N'" + year + "';";
            String SQL2 = "SELECT * FROM एमपीआर." + displayYojnaComboBox.getSelectedItem().toString() + " WHERE बांधकाम_विभाग = N'दक्षिण विभाग' AND वर्ष=N'" + year + "';";
            stmt = config.conn.createStatement();
            rs = stmt.executeQuery(SQL);

            setData(dataTable, rs);

            stmt2 = config.conn.createStatement();
            rs2 = stmt2.executeQuery(SQL2);
            setData(dataTable2, rs2);

            TABLE_COLUMNS.clear();
            TABLE_COLUMNS2.clear();
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(), "Monthly Progress Report", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();

        }

    }

    public void setData(JTable dataTable, ResultSet resultSet) {

        try {

            System.out.print("null available");
            ArrayList<ArrayList<Object>> result = toArrayList(resultSet);

            DefaultTableModel aModel = (DefaultTableModel) dataTable.getModel();
            aModel.getDataVector().removeAllElements();
            aModel.fireTableStructureChanged();
            System.out.println(result);
            Object[] object1 = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19};
            aModel.addRow(object1);
            for (int i = 0; i < result.size(); i++) {
                Object[] object = result.get(i).toArray();

                aModel.addRow(object);
                dataTable.setRowHeight(i, ROW_HEIGHT);

            }

            dataTable.setModel(aModel);

        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    public ArrayList<ArrayList<Object>> toArrayList(ResultSet resultSet) {
        ArrayList<ArrayList<Object>> table = null;
        try {

            int columnCount = resultSet.getMetaData().getColumnCount();
            if (resultSet.getType() == ResultSet.TYPE_FORWARD_ONLY) {
                table = new ArrayList<>();
            } else {

                resultSet.last();
                table = new ArrayList<>(resultSet.getRow());
                resultSet.beforeFirst();

            }

            for (ArrayList<Object> row; resultSet.next(); table.add(row)) {

                row = new ArrayList<>(columnCount);
                for (int c = 1; c <= columnCount; ++c) {

                    row.add(resultSet.getString(c).intern());

                }

            }

        } catch (Exception ex) {

        }

        return table;
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;

                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(AddData.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(AddData.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(AddData.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(AddData.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new AddData().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTable abstractTable1;
    private javax.swing.JTable abstractTable2;
    private javax.swing.JTextField akunTxt;
    private javax.swing.JTextField akunTxt1;
    private javax.swing.JButton allDataButton;
    private javax.swing.JButton backupBtn;
    private javax.swing.JButton closeBtn;
    private javax.swing.JButton closeBtn1;
    private javax.swing.JButton closeBtn2;
    private javax.swing.JButton closeBtn3;
    private javax.swing.JButton deleteBtn;
    private javax.swing.JComboBox<String> displayYojnaComboBox;
    private javax.swing.JButton editBtn;
    private javax.swing.JButton exportToExcel;
    private javax.swing.JButton exportToExcel1;
    private javax.swing.JButton exportToExcel2;
    private javax.swing.JButton exportToExcel3;
    private javax.swing.JTextField gstTxt;
    private javax.swing.JTextField gstTxt1;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton3;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel22;
    private javax.swing.JLabel jLabel23;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel25;
    private javax.swing.JLabel jLabel26;
    private javax.swing.JLabel jLabel27;
    private javax.swing.JLabel jLabel28;
    private javax.swing.JLabel jLabel29;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel30;
    private javax.swing.JLabel jLabel31;
    private javax.swing.JLabel jLabel32;
    private javax.swing.JLabel jLabel33;
    private javax.swing.JLabel jLabel34;
    private javax.swing.JLabel jLabel35;
    private javax.swing.JLabel jLabel36;
    private javax.swing.JLabel jLabel37;
    private javax.swing.JLabel jLabel38;
    private javax.swing.JLabel jLabel39;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JPanel jPanel7;
    private javax.swing.JPanel jPanel8;
    private javax.swing.JPanel jPanel9;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JTabbedPane jTabbedPane1;
    private com.toedter.calendar.JDateChooser kamacheAdeshDinank;
    private com.toedter.calendar.JDateChooser kamacheAdeshDinank1;
    private javax.swing.JComboBox<String> kamacheNaavComboBox;
    private javax.swing.JTextField kamacheNaavTxt;
    private com.toedter.calendar.JDateChooser kamachiMudatMahineDate;
    private com.toedter.calendar.JDateChooser kamachiMudatMahineDate1;
    private javax.swing.JTextField kharchWthGst;
    private javax.swing.JTextField kharchWthGstUpd;
    private javax.swing.JTextField maktedaracheNaavTxt;
    private javax.swing.JTextField maktedaracheNaavTxt1;
    private javax.swing.JTextField manyataRakkam;
    private javax.swing.JTextField manyataRakkam1;
    private javax.swing.JTextField nivedaRakkam;
    private javax.swing.JTextField nivedaRakkam1;
    private javax.swing.JTextField praptSun;
    private javax.swing.JTextField praptSunUpd;
    private com.toedter.calendar.JDateChooser prashashkiyaDinank;
    private com.toedter.calendar.JDateChooser prashashkiyaDinank1;
    private javax.swing.JButton saveBtn;
    public static javax.swing.JComboBox<String> sheraComboBox;
    private javax.swing.JComboBox<String> sheraComboBox1;
    private javax.swing.JButton showDataBtn;
    private javax.swing.JButton showMaktedarForm;
    private javax.swing.JComboBox<String> talukaComboBox;
    private javax.swing.JComboBox<String> talukaComboBox1;
    private com.toedter.calendar.JDateChooser tantrikDinank;
    private com.toedter.calendar.JDateChooser tantrikDinank1;
    private javax.swing.JTextField tantrikManyatRakkam;
    private javax.swing.JTextField tantrikManyatRakkam1;
    private javax.swing.JButton updateBtn;
    private javax.swing.JComboBox<String> updateyojnaComboBox;
    private javax.swing.JComboBox<String> vibhaagComboBox;
    private javax.swing.JComboBox<String> vibhaagComboBox1;
    private com.toedter.calendar.JDateChooser workCompleteDate;
    private com.toedter.calendar.JDateChooser workCompleteDate1;
    private javax.swing.JComboBox<String> yearChooseDisplay;
    private javax.swing.JComboBox<String> yearComboBox;
    private javax.swing.JComboBox<String> yearComboBox1;
    private javax.swing.JComboBox<String> yojnaComboBox;
    private javax.swing.JTable yojnaNorthTable;
    private javax.swing.JTable yojnaSouthTable;
    // End of variables declaration//GEN-END:variables
}

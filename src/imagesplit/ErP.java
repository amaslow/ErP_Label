package imagesplit;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.font.TextAttribute;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.AttributedString;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.awt.Rectangle;
import java.util.Iterator;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.Document;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

public class ErP extends javax.swing.JFrame {

    public ErP() {

        initComponents();
        jProgressBar1.setStringPainted(true);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {
        bindingGroup = new org.jdesktop.beansbinding.BindingGroup();

        ItemListbuttonGroup = new javax.swing.ButtonGroup();
        shownbuttonGroup = new javax.swing.ButtonGroup();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu2 = new javax.swing.JMenu();
        jMenu3 = new javax.swing.JMenu();
        ItemRadioButton = new javax.swing.JRadioButton();
        ItemLabel = new javax.swing.JLabel();
        ItemTextField = new javax.swing.JTextField();
        itemLabel1 = new javax.swing.JLabel();
        itemLabel2 = new javax.swing.JLabel();
        itemLabel3 = new javax.swing.JLabel();
        listRadioButton = new javax.swing.JRadioButton();
        listLabel = new javax.swing.JLabel();
        listTextField = new javax.swing.JTextField();
        listBrowseButton = new javax.swing.JButton();
        jSeparator1 = new javax.swing.JSeparator();
        shownLabel = new javax.swing.JLabel();
        shownItemLabel = new javax.swing.JLabel();
        shownItemRadioButton = new javax.swing.JRadioButton();
        shownSapLabel = new javax.swing.JLabel();
        shownSapRadioButton = new javax.swing.JRadioButton();
        shownEanLabel = new javax.swing.JLabel();
        shownEanRadioButton = new javax.swing.JRadioButton();
        jSeparator2 = new javax.swing.JSeparator();
        lanLabel = new javax.swing.JLabel();
        lanAllCheckBox = new javax.swing.JCheckBox();
        lanGBCheckBox = new javax.swing.JCheckBox();
        lanDECheckBox = new javax.swing.JCheckBox();
        lanFRCheckBox = new javax.swing.JCheckBox();
        lanNLCheckBox = new javax.swing.JCheckBox();
        lanBACheckBox = new javax.swing.JCheckBox();
        lanBGCheckBox = new javax.swing.JCheckBox();
        lanCZCheckBox = new javax.swing.JCheckBox();
        lanDKCheckBox = new javax.swing.JCheckBox();
        lanEECheckBox = new javax.swing.JCheckBox();
        lanESCheckBox = new javax.swing.JCheckBox();
        lanFICheckBox = new javax.swing.JCheckBox();
        lanGRCheckBox = new javax.swing.JCheckBox();
        lanHRCheckBox = new javax.swing.JCheckBox();
        lanHUCheckBox = new javax.swing.JCheckBox();
        lanISCheckBox = new javax.swing.JCheckBox();
        lanITCheckBox = new javax.swing.JCheckBox();
        lanLTCheckBox = new javax.swing.JCheckBox();
        lanLVCheckBox = new javax.swing.JCheckBox();
        lanPLCheckBox = new javax.swing.JCheckBox();
        lanPTCheckBox = new javax.swing.JCheckBox();
        lanROCheckBox = new javax.swing.JCheckBox();
        lanRSCheckBox = new javax.swing.JCheckBox();
        lanRUCheckBox = new javax.swing.JCheckBox();
        lanSECheckBox = new javax.swing.JCheckBox();
        lanSICheckBox = new javax.swing.JCheckBox();
        lanSKCheckBox = new javax.swing.JCheckBox();
        lanTRCheckBox = new javax.swing.JCheckBox();
        lanUACheckBox = new javax.swing.JCheckBox();
        jSeparator3 = new javax.swing.JSeparator();
        orientLabel = new javax.swing.JLabel();
        orientPCheckBox = new javax.swing.JCheckBox();
        orientLCheckBox = new javax.swing.JCheckBox();
        jSeparator4 = new javax.swing.JSeparator();
        outLabel = new javax.swing.JLabel();
        outTextCheckBox = new javax.swing.JCheckBox();
        outTextField = new javax.swing.JTextField();
        jSeparator5 = new javax.swing.JSeparator();
        outLangCheckBox = new javax.swing.JCheckBox();
        jSeparator6 = new javax.swing.JSeparator();
        outItemComboBox = new javax.swing.JComboBox();
        jSeparator7 = new javax.swing.JSeparator();
        outOrientCheckBox = new javax.swing.JCheckBox();
        jSeparator8 = new javax.swing.JSeparator();
        outDateCheckBox = new javax.swing.JCheckBox();
        outDateChooser = new com.toedter.calendar.JDateChooser();
        jSeparator9 = new javax.swing.JSeparator();
        outExtComboBox = new javax.swing.JComboBox();
        outPdfCheckBox = new javax.swing.JCheckBox();
        outExampleLabel = new javax.swing.JLabel();
        listStartButton = new javax.swing.JButton();
        jProgressBar1 = new javax.swing.JProgressBar();
        jRowCounterLabel = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        oneFolderCheckBox = new javax.swing.JCheckBox();

        jMenu2.setText("File");
        jMenuBar1.add(jMenu2);

        jMenu3.setText("Edit");
        jMenuBar1.add(jMenu3);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("ErP Luminaire Labeling");
        setMinimumSize(new java.awt.Dimension(420, 484));
        setPreferredSize(new java.awt.Dimension(430, 550));
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        ItemListbuttonGroup.add(ItemRadioButton);
        ItemRadioButton.setSelected(true);
        getContentPane().add(ItemRadioButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 30, -1, 20));

        ItemLabel.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        ItemLabel.setText("1 item:");

        org.jdesktop.beansbinding.Binding binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, ItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), ItemLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(ItemLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 30, 50, 20));

        ItemTextField.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        ItemTextField.setToolTipText("Enter here either ITEM number or SAP number or EAN code");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, ItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), ItemTextField, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(ItemTextField, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 18, 170, 40));
        ItemTextField.getAccessibleContext().setAccessibleDescription("");
        ItemTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void changedUpdate(DocumentEvent e) {
                changed();
            }
            public void removeUpdate(DocumentEvent e) {
                changed();
            }
            public void insertUpdate(DocumentEvent e) {
                changed();
            }
            public void changed() {
                if (!ItemTextField.getText().equals("") && ItemRadioButton.isSelected()){
                    listStartButton.setEnabled(true);
                }
                else {
                    listStartButton.setEnabled(false);
                }
            }
        });

        itemLabel1.setFont(new java.awt.Font("Tahoma", 2, 10)); // NOI18N
        itemLabel1.setText("- ITEM number");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, ItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), itemLabel1, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(itemLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 20, -1, 13));

        itemLabel2.setFont(new java.awt.Font("Tahoma", 2, 10)); // NOI18N
        itemLabel2.setText("- SAP number");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, ItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), itemLabel2, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(itemLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 33, -1, 13));

        itemLabel3.setFont(new java.awt.Font("Tahoma", 2, 10)); // NOI18N
        itemLabel3.setText("- EAN code");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, ItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), itemLabel3, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(itemLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 46, 60, 13));

        ItemListbuttonGroup.add(listRadioButton);
        getContentPane().add(listRadioButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 63, -1, 30));

        listLabel.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        listLabel.setText("List of items:");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, listRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), listLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(listLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 63, 80, 30));

        listTextField.setFont(new java.awt.Font("Tahoma", 1, 8)); // NOI18N
        listTextField.setToolTipText("Path to Excel list with items");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, listRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), listTextField, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(listTextField, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 63, 215, 30));
        listTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void changedUpdate(DocumentEvent e) {
                changed();
            }
            public void removeUpdate(DocumentEvent e) {
                changed();
            }
            public void insertUpdate(DocumentEvent e) {
                changed();
            }
            public void changed() {
                if (!listTextField.getText().equals("") && listRadioButton.isSelected()){
                    listStartButton.setEnabled(true);
                }
                else {
                    listStartButton.setEnabled(false);
                }
            }
        });

        listBrowseButton.setText("Browse");
        listBrowseButton.setToolTipText("Click to browse Excel list with items");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, listRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), listBrowseButton, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        listBrowseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                listBrowseButtonActionPerformed(evt);
            }
        });
        getContentPane().add(listBrowseButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(330, 65, 70, 30));
        getContentPane().add(jSeparator1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 110, 400, 10));

        shownLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownLabel.setText("Item shown on the label as:");
        getContentPane().add(shownLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 120, -1, 36));

        shownItemLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownItemLabel.setText("Item");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, shownItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), shownItemLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(shownItemLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(190, 120, -1, -1));

        shownbuttonGroup.add(shownItemRadioButton);
        shownItemRadioButton.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownItemRadioButton.setSelected(true);
        shownItemRadioButton.setToolTipText("ITEM number will be shown on the label");
        shownItemRadioButton.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        shownItemRadioButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        shownItemRadioButton.setName("Item"); // NOI18N
        shownItemRadioButton.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        shownItemRadioButton.setVerticalTextPosition(javax.swing.SwingConstants.TOP);
        getContentPane().add(shownItemRadioButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(190, 130, -1, 20));

        shownSapLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownSapLabel.setText("SAP");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, shownSapRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), shownSapLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(shownSapLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(230, 120, -1, -1));

        shownbuttonGroup.add(shownSapRadioButton);
        shownSapRadioButton.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownSapRadioButton.setToolTipText("SAP number will be shown on the label");
        shownSapRadioButton.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        shownSapRadioButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        shownSapRadioButton.setName("SAP"); // NOI18N
        shownSapRadioButton.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        shownSapRadioButton.setVerticalTextPosition(javax.swing.SwingConstants.TOP);
        getContentPane().add(shownSapRadioButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(230, 130, -1, 20));

        shownEanLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownEanLabel.setText("EAN");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, shownEanRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), shownEanLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(shownEanLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(270, 120, -1, -1));

        shownbuttonGroup.add(shownEanRadioButton);
        shownEanRadioButton.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        shownEanRadioButton.setToolTipText("EAN code will be shown on the label");
        shownEanRadioButton.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        shownEanRadioButton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        shownEanRadioButton.setName("EAN"); // NOI18N
        shownEanRadioButton.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        shownEanRadioButton.setVerticalTextPosition(javax.swing.SwingConstants.TOP);
        getContentPane().add(shownEanRadioButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(270, 130, -1, 20));
        getContentPane().add(jSeparator2, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 155, 400, 10));

        lanLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        lanLabel.setText("Languages:");
        getContentPane().add(lanLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 160, -1, 21));

        lanAllCheckBox.setFont(new java.awt.Font("Tahoma", 2, 10)); // NOI18N
        lanAllCheckBox.setText("All on/off");
        lanAllCheckBox.setToolTipText("Choose all languages as on or off");
        lanAllCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanAllCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        lanAllCheckBox.setPreferredSize(new java.awt.Dimension(36, 21));
        lanAllCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                lanAllCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(lanAllCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(85, 160, 67, -1));

        lanGBCheckBox.setSelected(true);
        lanGBCheckBox.setText("EN");
        lanGBCheckBox.setToolTipText("English");
        lanGBCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanGBCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanGBCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(210, 170, 40, -1));

        lanDECheckBox.setSelected(true);
        lanDECheckBox.setText("DE");
        lanDECheckBox.setToolTipText("German");
        lanDECheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanDECheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanDECheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(260, 170, 40, -1));

        lanFRCheckBox.setSelected(true);
        lanFRCheckBox.setText("FR");
        lanFRCheckBox.setToolTipText("French");
        lanFRCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanFRCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanFRCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 170, 40, -1));

        lanNLCheckBox.setSelected(true);
        lanNLCheckBox.setText("NL");
        lanNLCheckBox.setToolTipText("Dutch");
        lanNLCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanNLCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanNLCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(360, 170, 40, -1));

        lanBACheckBox.setText("BA");
        lanBACheckBox.setToolTipText("Bosnian and Hercegovina");
        lanBACheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanBACheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanBACheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 190, 40, -1));

        lanBGCheckBox.setText("BG");
        lanBGCheckBox.setToolTipText("Bulgarian");
        lanBGCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanBGCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanBGCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(60, 190, 40, -1));

        lanCZCheckBox.setText("CZ");
        lanCZCheckBox.setToolTipText("Czech");
        lanCZCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanCZCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanCZCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 190, 40, -1));

        lanDKCheckBox.setText("DK");
        lanDKCheckBox.setToolTipText("Danish");
        lanDKCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanDKCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanDKCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(160, 190, 40, -1));

        lanEECheckBox.setText("EE");
        lanEECheckBox.setToolTipText("Estonian");
        lanEECheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanEECheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanEECheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(210, 190, 40, -1));

        lanESCheckBox.setText("ES");
        lanESCheckBox.setToolTipText("Spanish");
        lanESCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanESCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanESCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(260, 190, 40, -1));

        lanFICheckBox.setText("FI");
        lanFICheckBox.setToolTipText("Finnish");
        lanFICheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanFICheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanFICheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 190, 40, -1));

        lanGRCheckBox.setText("GR");
        lanGRCheckBox.setToolTipText("Greek");
        lanGRCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanGRCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanGRCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(360, 190, 40, -1));

        lanHRCheckBox.setText("HR");
        lanHRCheckBox.setToolTipText("Croatian");
        lanHRCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanHRCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanHRCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 210, 40, -1));

        lanHUCheckBox.setText("HU");
        lanHUCheckBox.setToolTipText("Hungarian");
        lanHUCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanHUCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanHUCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(60, 210, 40, -1));

        lanISCheckBox.setText("IS");
        lanISCheckBox.setToolTipText("Icelandic");
        lanISCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanISCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanISCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 210, 40, -1));

        lanITCheckBox.setText("IT");
        lanITCheckBox.setToolTipText("Italian");
        lanITCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanITCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanITCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(160, 210, 40, -1));

        lanLTCheckBox.setText("LT");
        lanLTCheckBox.setToolTipText("Lithuanian");
        lanLTCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanLTCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanLTCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(210, 210, 40, -1));

        lanLVCheckBox.setText("LV");
        lanLVCheckBox.setToolTipText("Latvian");
        lanLVCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanLVCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanLVCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(260, 210, 40, -1));

        lanPLCheckBox.setText("PL");
        lanPLCheckBox.setToolTipText("Polish");
        lanPLCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanPLCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanPLCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 210, 40, -1));

        lanPTCheckBox.setText("PT");
        lanPTCheckBox.setToolTipText("Portuguese");
        lanPTCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanPTCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanPTCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(360, 210, 40, -1));

        lanROCheckBox.setText("RO");
        lanROCheckBox.setToolTipText("Romanian");
        lanROCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanROCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanROCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 230, 40, -1));

        lanRSCheckBox.setText("RS");
        lanRSCheckBox.setToolTipText("Serbian");
        lanRSCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanRSCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanRSCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(60, 230, 40, -1));

        lanRUCheckBox.setText("RU");
        lanRUCheckBox.setToolTipText("Russian");
        lanRUCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanRUCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanRUCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(110, 230, 40, -1));

        lanSECheckBox.setText("SE");
        lanSECheckBox.setToolTipText("Swedish");
        lanSECheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanSECheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanSECheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(160, 230, 40, -1));

        lanSICheckBox.setText("SI");
        lanSICheckBox.setToolTipText("Slovenian");
        lanSICheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanSICheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanSICheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(210, 230, 40, -1));

        lanSKCheckBox.setText("SK");
        lanSKCheckBox.setToolTipText("Slovak");
        lanSKCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanSKCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanSKCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(260, 230, 40, -1));

        lanTRCheckBox.setText("TR");
        lanTRCheckBox.setToolTipText("Turkish");
        lanTRCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanTRCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanTRCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 230, 40, -1));

        lanUACheckBox.setText("UA");
        lanUACheckBox.setToolTipText("Ukrainian");
        lanUACheckBox.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        lanUACheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(lanUACheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(360, 230, 40, -1));
        getContentPane().add(jSeparator3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 260, 400, 10));

        orientLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        orientLabel.setText("Orientation:");
        getContentPane().add(orientLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 270, -1, 21));

        orientPCheckBox.setSelected(true);
        orientPCheckBox.setText("Portrait");
        orientPCheckBox.setToolTipText("portrait orientation");
        orientPCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        orientPCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                orientPCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(orientPCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(140, 270, -1, -1));

        orientLCheckBox.setSelected(true);
        orientLCheckBox.setText("Landscape");
        orientLCheckBox.setToolTipText("landscape orientation");
        orientLCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        orientLCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                orientLCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(orientLCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(230, 270, -1, -1));
        getContentPane().add(jSeparator4, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 300, 400, 10));

        outLabel.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        outLabel.setText("Filename:");
        getContentPane().add(outLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 310, -1, 21));

        outTextCheckBox.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        outTextCheckBox.setSelected(true);
        outTextCheckBox.setText("text");
        outTextCheckBox.setToolTipText("add text info in filename");
        outTextCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        outTextCheckBox.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        outTextCheckBox.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        outTextCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outTextCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outTextCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 340, 60, 20));

        outTextField.setFont(new java.awt.Font("Tahoma", 0, 10)); // NOI18N
        outTextField.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        outTextField.setText("Energylabel");
        outTextField.setToolTipText("type text info in filename");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, outTextCheckBox, org.jdesktop.beansbinding.ELProperty.create("${selected}"), outTextField, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(outTextField, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 360, 70, 23));

        jSeparator5.setOrientation(javax.swing.SwingConstants.VERTICAL);
        getContentPane().add(jSeparator5, new org.netbeans.lib.awtextra.AbsoluteConstraints(81, 340, 5, 50));

        outLangCheckBox.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        outLangCheckBox.setSelected(true);
        outLangCheckBox.setText("language");
        outLangCheckBox.setToolTipText("add info about language");
        outLangCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        outLangCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        outLangCheckBox.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        outLangCheckBox.setVerticalTextPosition(javax.swing.SwingConstants.TOP);
        outLangCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outLangCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outLangCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(80, 340, 45, 40));

        jSeparator6.setOrientation(javax.swing.SwingConstants.VERTICAL);
        getContentPane().add(jSeparator6, new org.netbeans.lib.awtextra.AbsoluteConstraints(121, 340, 5, 50));

        outItemComboBox.setMaximumRowCount(3);
        outItemComboBox.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Item", "SAP", "EAN" }));
        outItemComboBox.setSelectedIndex(1);
        outItemComboBox.setToolTipText("choose description between Item number, SAP number and EAN code");
        outItemComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outItemComboBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outItemComboBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(127, 360, -1, -1));

        jSeparator7.setOrientation(javax.swing.SwingConstants.VERTICAL);
        getContentPane().add(jSeparator7, new org.netbeans.lib.awtextra.AbsoluteConstraints(186, 340, 5, 50));

        outOrientCheckBox.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        outOrientCheckBox.setSelected(true);
        outOrientCheckBox.setText("orientation");
        outOrientCheckBox.setToolTipText("add info about orientation");
        outOrientCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        outOrientCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        outOrientCheckBox.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        outOrientCheckBox.setVerticalTextPosition(javax.swing.SwingConstants.TOP);
        outOrientCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outOrientCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outOrientCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(185, 340, 51, 40));

        jSeparator8.setOrientation(javax.swing.SwingConstants.VERTICAL);
        getContentPane().add(jSeparator8, new org.netbeans.lib.awtextra.AbsoluteConstraints(232, 340, 5, 50));

        outDateCheckBox.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        outDateCheckBox.setSelected(true);
        outDateCheckBox.setText("date");
        outDateCheckBox.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        outDateCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        outDateCheckBox.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        outDateCheckBox.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        outDateCheckBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outDateCheckBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outDateCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(235, 340, 70, -1));

        outDateChooser.setToolTipText("addd info about the date");
        outDateChooser.setDate(Calendar.getInstance().getTime());
        outDateChooser.setDateFormatString("yyyyMMdd");
        outDateChooser.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, outDateCheckBox, org.jdesktop.beansbinding.ELProperty.create("${selected}"), outDateChooser, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        getContentPane().add(outDateChooser, new org.netbeans.lib.awtextra.AbsoluteConstraints(235, 360, 100, 23));

        jSeparator9.setOrientation(javax.swing.SwingConstants.VERTICAL);
        getContentPane().add(jSeparator9, new org.netbeans.lib.awtextra.AbsoluteConstraints(336, 340, 5, 50));

        outExtComboBox.setMaximumRowCount(2);
        outExtComboBox.setModel(new javax.swing.DefaultComboBoxModel(new String[] { ".png", ".jpg" }));
        outExtComboBox.setToolTipText("choose file extension");
        outExtComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                outExtComboBoxActionPerformed(evt);
            }
        });
        getContentPane().add(outExtComboBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(340, 360, -1, -1));

        outPdfCheckBox.setFont(new java.awt.Font("Tahoma", 1, 10)); // NOI18N
        outPdfCheckBox.setForeground(new java.awt.Color(0, 153, 51));
        outPdfCheckBox.setText("+ PDF");
        outPdfCheckBox.setToolTipText("make extra labels in  pdf format");
        outPdfCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(outPdfCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(330, 390, -1, -1));

        outExampleLabel.setFont(new java.awt.Font("Tahoma", 2, 12)); // NOI18N
        outExampleLabel.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        outExampleLabel.setText("Example:  Energylabel_EN_1003676_P_20150101.png");
        getContentPane().add(outExampleLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 410, 380, -1));

        listStartButton.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        listStartButton.setText("START");
        listStartButton.setEnabled(false);
        listStartButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                listStartButtonActionPerformed(evt);
            }
        });
        getContentPane().add(listStartButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 440, 90, 40));
        getContentPane().add(jProgressBar1, new org.netbeans.lib.awtextra.AbsoluteConstraints(180, 470, 215, -1));
        getContentPane().add(jRowCounterLabel, new org.netbeans.lib.awtextra.AbsoluteConstraints(170, 480, -1, -1));

        jLabel1.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        jLabel1.setText("extension");
        jLabel1.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        getContentPane().add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(342, 340, -1, 18));

        jLabel2.setFont(new java.awt.Font("Tahoma", 0, 9)); // NOI18N
        jLabel2.setText("description");
        jLabel2.setVerticalAlignment(javax.swing.SwingConstants.BOTTOM);
        getContentPane().add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(132, 340, -1, 18));

        oneFolderCheckBox.setFont(new java.awt.Font("Tahoma", 0, 10)); // NOI18N
        oneFolderCheckBox.setText("(all files into 1 folder)");
        oneFolderCheckBox.setHorizontalTextPosition(javax.swing.SwingConstants.LEFT);
        getContentPane().add(oneFolderCheckBox, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 310, -1, -1));

        bindingGroup.bind();

        pack();
    }// </editor-fold>//GEN-END:initComponents
        String mainfolder = "G:\\Share Company Wide\\Company Transfer\\ERP classificatie";
        String productContent = "G:\\Product Content\\PRODUCTS\\";

    private int findRow(XSSFSheet sheet, String item) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(item)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return 0;
    }

    private void label(String item, JFileChooser dest,List<String> noitem) throws IOException {
        
        String item1 = item.replace("/", "_");
        String subfolder = item1 + "\\";
        if(oneFolderCheckBox.isSelected()){
            subfolder = "";
        }
        File subdir = new File(dest.getSelectedFile() + "\\" + subfolder);


        File sources = new File(mainfolder + "\\ERP_Elements");
        String excelname = mainfolder + "\\ERP.xlsx";

        JCheckBox[] orientCheckBoxes = {orientPCheckBox, orientLCheckBox};

        ArrayList<String> orient = new ArrayList<String>();
        for (int i = 0; i < orientCheckBoxes.length; i += 1) {
            if (orientCheckBoxes[i].isSelected()) {
                orient.add(orientCheckBoxes[i].getText());
            }
        }
        JCheckBox[] langCheckBoxes = {lanBACheckBox, lanBGCheckBox, lanCZCheckBox, lanDECheckBox, lanDKCheckBox,
            lanEECheckBox, lanESCheckBox, lanFICheckBox, lanFRCheckBox, lanGBCheckBox, lanGRCheckBox,
            lanHRCheckBox, lanHUCheckBox, lanISCheckBox, lanITCheckBox, lanLTCheckBox, lanLVCheckBox,
            lanNLCheckBox, lanPLCheckBox, lanPTCheckBox, lanROCheckBox, lanRSCheckBox, lanRUCheckBox,
            lanSECheckBox, lanSICheckBox, lanSKCheckBox, lanTRCheckBox, lanUACheckBox};
        ArrayList<String> lang = new ArrayList<String>();
        for (int i = 0; i < langCheckBoxes.length; i += 1) {
            if (langCheckBoxes[i].isSelected()) {
                lang.add(langCheckBoxes[i].getText());
            }
        }

        FileInputStream fis = null;
        fis = new FileInputStream(excelname);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        int rownr = findRow(sheet, item);

        if (rownr != 0) {
            for (int o = 0; o < orient.size(); o++) {
                for (int l = 0; l < lang.size(); l++) {
                    jProgressBar1.setValue(l * 7);

                    XSSFRow row = sheet.getRow(rownr);

                    XSSFCell itemNo1 = row.getCell(0);// get item number
                    String itemNo = itemNo1.getStringCellValue();

                    XSSFCell sap1 = row.getCell(1); // get sap number
                    String sap = sap1.getStringCellValue();
                    String sapNodot = sap.replace(".", "");

                    XSSFCell ean1 = row.getCell(2); // get ean code
                    String ean = ean1.getStringCellValue();

                    XSSFCell en = row.getCell(3); // get energy logo 

                    XSSFCell t_b_p = row.getCell(4, org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK); //get text bottom
                    switch (t_b_p.getCellType()) {
                        case XSSFCell.CELL_TYPE_BLANK:
                            t_b_p.setCellValue("Text_bottom_00");
                            break;
                    }
                    String tbp_n = t_b_p.getStringCellValue().substring(12);
                    String tbl_n = Integer.toString(Integer.parseInt(tbp_n) + 15);
                    switch (tbl_n) {
                        case "30":
                            tbl_n = "15";
                            break;
                        case "15":
                            tbl_n = "00";
                            break;
                    }

                    XSSFCell t_t = row.getCell(5); //get text top
                    String tt_n = t_t.getStringCellValue().substring(9);

                    XSSFCell log = row.getCell(6); //get logo

                    XSSFCell icon = row.getCell(7); //get icon (indoor/outdoor)

                    String text = "";
                    if (outTextCheckBox.isSelected()) {
                        text = outTextField.getText() + "_";
                    }
                    String language = "";
                    if (outLangCheckBox.isSelected()) {
                        language = lang.get(l) + "_";
                    } else {
                        language = "";
                    }
                    String description = "";
                    switch (outItemComboBox.getSelectedItem().toString()) {
                        case "Item":
                            description = item1;
                            break;
                        case "SAP":
                            description = sapNodot;
                            break;
                        case "EAN":
                            description = ean;
                            break;
                    }
                    String orientation = "";
                    if (outOrientCheckBox.isSelected()) {
                        orientation = "_" + orient.get(o).substring(0, 1);
                    } else {
                        orientation = "";
                    }
                    String outDate = "";
                    if (outDateCheckBox.isSelected()) {
                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
                        outDate = "_" + dateFormat.format(outDateChooser.getDate());
                    } else {
                        outDate = "";
                    }
                    String ext = (String) outExtComboBox.getSelectedItem().toString();
                    File output = new File(subdir + "\\" + text + language + description + outDate + orientation + ext);

                    BufferedImage base = ImageIO.read(new File(sources + "\\Base_" + orient.get(o) + ".jpg"));
                    BufferedImage logo = ImageIO.read(new File(sources + "\\" + log + ".png"));
                    BufferedImage icon_indoor = ImageIO.read(new File(sources + "\\" + icon + ".png"));
                    BufferedImage text_top = ImageIO.read(new File(sources + "\\taal_" + lang.get(l) + "\\Text_top_" + lang.get(l) + "_" + tt_n + ".png"));
                    BufferedImage energy = ImageIO.read(new File(sources + "\\" + en + ".png"));
                    BufferedImage text_bottom_p = ImageIO.read(new File(sources + "\\taal_" + lang.get(l) + "\\Text_bottom_" + lang.get(l) + "_" + tbp_n + ".png"));
                    BufferedImage text_bottom_l = ImageIO.read(new File(sources + "\\taal_" + lang.get(l) + "\\Text_bottom_" + lang.get(l) + "_" + tbl_n + ".png"));

                    // create the new image, canvas size is the size of the base image
                    int w = base.getWidth();
                    int h = base.getHeight();
                    BufferedImage combined = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
                    Graphics g = combined.getGraphics();

                    //g.setFont(g.getFont().deriveFont(30f));
                    String shown1 = new GroupButtonUtils().getSelectedButtonName(shownbuttonGroup);
                    String shown = "";
                    switch (shown1) {
                        case "Item":
                            shown = itemNo;
                            break;
                        case "SAP":
                            shown = sap;
                            break;
                        case "EAN":
                            shown = ean;
                            break;
                    }
                    AttributedString word = new AttributedString(shown);
                    int item_l = shown.length();
                    if (item_l < 11) {
                        word.addAttribute(TextAttribute.FONT, new Font("Calibri", Font.BOLD, 40));
                        word.addAttribute(TextAttribute.FOREGROUND, Color.BLACK);
                    } else {
                        word.addAttribute(TextAttribute.FONT, new Font("Calibri", Font.BOLD, 25));
                        word.addAttribute(TextAttribute.FOREGROUND, Color.BLACK);
                    }

                    //combined.createGraphics().drawImage(combined, 0, 0, Color.YELLOW, null);
                    // paint both images, preserving the alpha channels
                    g.drawImage(base, 0, 0, null);

                    switch (orient.get(o)) {
                        case "Portrait":
                            g.drawImage(logo, 20, 45, null);
                            g.drawString(word.getIterator(), 370, 135);
                            g.drawImage(icon_indoor, 40, 200, null);
                            g.drawImage(text_top, 220, 200, null);
                            g.drawImage(energy, 40, 400, null);
                            g.drawImage(text_bottom_p, 35, 900, null);
                            break;
                        case "Landscape":
                            g.drawImage(logo, 20, 445, null);
                            g.drawString(word.getIterator(), 370, 530);
                            g.drawImage(icon_indoor, 40, 45, null);
                            g.drawImage(text_top, 220, 45, null);
                            g.drawImage(energy, 590, 45, null);
                            g.drawImage(text_bottom_l, 40, 220, null);
                            break;
                    }

                    if (o < 1) {
                        int progress = (int) (((l + 1) * 1.79) / 2);
                        jProgressBar1.setValue(progress);
                        Rectangle progressRect = jProgressBar1.getBounds();
                        progressRect.x = 0;
                        progressRect.y = 0;
                        jProgressBar1.paintImmediately(progressRect);

                    } else {
                        int progress = (int) (((l + 1) * 1.79) + 50);
                        jProgressBar1.setValue(progress);
                        Rectangle progressRect = jProgressBar1.getBounds();
                        progressRect.x = 0;
                        progressRect.y = 0;
                        jProgressBar1.paintImmediately(progressRect);
                    }
                    g.dispose();

                    // Save as new image

                    if (!subdir.exists()) {
                        subdir.mkdir();
                    }
                    
                    ImageIO.write(combined, "PNG", output);
                    if (outPdfCheckBox.isSelected()) {
                        try {
                            File pdfDir = new File(subdir + "\\PDF\\");
                            pdfDir.mkdir();
                            String pdf = subdir + "\\PDF\\" + text + language + description + outDate + orientation + ".pdf";

                            Image image = Image.getInstance(output.toString());
                            image.scalePercent((float) 24.01);
                            float image_w = image.getScaledWidth();
                            float image_h = image.getScaledHeight();
                            com.itextpdf.text.Rectangle rect = new com.itextpdf.text.Rectangle(image_w, image_h);
                            Document document = new Document();
                            document.setPageSize(rect);
                            FileOutputStream fos = new FileOutputStream(pdf);
                            PdfWriter writer = PdfWriter.getInstance(document, fos);
                            writer.open();
                            document.open();
                            image.setAbsolutePosition(0, 0);
                            document.add(image);
                            document.close();
                            writer.close();
                        } catch (Exception i1) {
                            i1.printStackTrace();
                        }
                    }
                }
            }

        } else {
            //JOptionPane.showMessageDialog(null, "There is no data for item " + item);
            noitem.add(item);
        }
    }

    private void listBrowseButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_listBrowseButtonActionPerformed
        JFileChooser list = new JFileChooser(mainfolder);
        list.setDialogTitle("Select excel file with list");
        list.setFileSelectionMode(JFileChooser.FILES_ONLY);
        list.showOpenDialog(null);

        listTextField.setText(list.getSelectedFile().getPath());
    }//GEN-LAST:event_listBrowseButtonActionPerformed

    private void listStartButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_listStartButtonActionPerformed
        List<String> noitem = new ArrayList<String>();
        try {
            JFileChooser dest = new JFileChooser(mainfolder);
            dest.setDialogTitle("Select destination folder");
            dest.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            dest.showSaveDialog(null);

            if (listRadioButton.isSelected()) {
                String path = listTextField.getText();
                FileInputStream fis1 = null;
                fis1 = new FileInputStream(path);
                XSSFWorkbook wb = new XSSFWorkbook(fis1);
                XSSFSheet sheet = wb.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    
                    Cell cell = cellIterator.next();
                    String item = cell.getStringCellValue();
                    label(item, dest,noitem);
                }
                if (noitem.size() > 0) {
                JOptionPane.showMessageDialog(null, "Generator didn't find data for folowing items: " + noitem.toString());
                }
                File subdir = new File(dest.getSelectedFile() + "\\");
                Desktop desktop = Desktop.getDesktop();
                desktop.open(subdir);
            } else {
                String item = ItemTextField.getText().toUpperCase();
                label(item, dest,noitem);
                String item1 = item.replace("/", "_");
                File subdir = new File(dest.getSelectedFile() + "\\" + item1 + "\\");
                Desktop desktop = Desktop.getDesktop();
                desktop.open(subdir);
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_listStartButtonActionPerformed

    private void lanAllCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_lanAllCheckBoxActionPerformed
        JCheckBox[] checkBoxes = {lanBACheckBox, lanBGCheckBox, lanCZCheckBox, lanDECheckBox, lanDKCheckBox,
            lanEECheckBox, lanESCheckBox, lanFICheckBox, lanFRCheckBox, lanGBCheckBox, lanGRCheckBox,
            lanHRCheckBox, lanHUCheckBox, lanISCheckBox, lanITCheckBox, lanLTCheckBox, lanLVCheckBox,
            lanNLCheckBox, lanPLCheckBox, lanPTCheckBox, lanROCheckBox, lanRSCheckBox, lanRUCheckBox,
            lanSECheckBox, lanSICheckBox, lanSKCheckBox, lanTRCheckBox, lanUACheckBox};
        if (!lanAllCheckBox.isSelected()) {
            for (int i = 0; i < checkBoxes.length; i = i + 1) {
                checkBoxes[i].setSelected(false);
            }
        } else {
            for (int i = 0; i < checkBoxes.length; i = i + 1) {
                checkBoxes[i].setSelected(true);
            }
        }

    }//GEN-LAST:event_lanAllCheckBoxActionPerformed

    private void orientPCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_orientPCheckBoxActionPerformed
        if (!orientPCheckBox.isSelected()) {
            orientLCheckBox.setSelected(true);
        }
        if (orientPCheckBox.isSelected() && orientLCheckBox.isSelected()) {
            outOrientCheckBox.setSelected(true);
        }
    }//GEN-LAST:event_orientPCheckBoxActionPerformed

    private void orientLCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_orientLCheckBoxActionPerformed
        if (!orientLCheckBox.isSelected()) {
            orientPCheckBox.setSelected(true);
        }
        if (orientPCheckBox.isSelected() && orientLCheckBox.isSelected()) {
            outOrientCheckBox.setSelected(true);
        }
     }//GEN-LAST:event_orientLCheckBoxActionPerformed

    private void outOrientCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outOrientCheckBoxActionPerformed
        if (orientPCheckBox.isSelected() && orientLCheckBox.isSelected()) {
            outOrientCheckBox.setSelected(true);
        }
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outOrientCheckBoxActionPerformed

    private void outLangCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outLangCheckBoxActionPerformed
        ArrayList<String> lang = new ArrayList<String>();
        lang.clear();
        JCheckBox[] langCheckBoxes = {lanBACheckBox, lanBGCheckBox, lanCZCheckBox, lanDECheckBox, lanDKCheckBox,
            lanEECheckBox, lanESCheckBox, lanFICheckBox, lanFRCheckBox, lanGBCheckBox, lanGRCheckBox,
            lanHRCheckBox, lanHUCheckBox, lanISCheckBox, lanITCheckBox, lanLTCheckBox, lanLVCheckBox,
            lanNLCheckBox, lanPLCheckBox, lanPTCheckBox, lanROCheckBox, lanRSCheckBox, lanRUCheckBox,
            lanSECheckBox, lanSICheckBox, lanSKCheckBox, lanTRCheckBox, lanUACheckBox};

        for (int i = 0; i < langCheckBoxes.length; i += 1) {
            if (langCheckBoxes[i].isSelected()) {
                lang.add("1");
            }
        }
        if (lang.size() > 1) {
            outLangCheckBox.setSelected(true);
        }
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outLangCheckBoxActionPerformed

    private void outItemComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outItemComboBoxActionPerformed
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outItemComboBoxActionPerformed

    private void outExtComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outExtComboBoxActionPerformed
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outExtComboBoxActionPerformed

    private void outTextCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outTextCheckBoxActionPerformed
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outTextCheckBoxActionPerformed

    private void outDateCheckBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_outDateCheckBoxActionPerformed
        String text = "";
        if (outTextCheckBox.isSelected()) {
            text = "Energylabel_";
        }
        String language = "";
        if (outLangCheckBox.isSelected()) {
            language = "EN_";
        }
        String description = "";
        switch (outItemComboBox.getSelectedIndex()) {
            case 0:
                description = "HL120";
                break;
            case 1:
                description = "1003676";
                break;
            case 2:
                description = "8711658257747";
                break;
        }
        String orientation = "";
        if (outOrientCheckBox.isSelected()) {
            orientation = "_P";
        }
        String outDate = "";
        if (outDateCheckBox.isSelected()) {
            outDate = "_20150101";
        }
        String extension = (String) outExtComboBox.getSelectedItem().toString();
        outExampleLabel.setText("Example:   " + text + language + description + orientation + outDate + extension);
    }//GEN-LAST:event_outDateCheckBoxActionPerformed

    /**
     *
     * @param args
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
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ErP().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel ItemLabel;
    private javax.swing.ButtonGroup ItemListbuttonGroup;
    private javax.swing.JRadioButton ItemRadioButton;
    private javax.swing.JTextField ItemTextField;
    private javax.swing.JLabel itemLabel1;
    private javax.swing.JLabel itemLabel2;
    private javax.swing.JLabel itemLabel3;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JLabel jRowCounterLabel;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JSeparator jSeparator3;
    private javax.swing.JSeparator jSeparator4;
    private javax.swing.JSeparator jSeparator5;
    private javax.swing.JSeparator jSeparator6;
    private javax.swing.JSeparator jSeparator7;
    private javax.swing.JSeparator jSeparator8;
    private javax.swing.JSeparator jSeparator9;
    private javax.swing.JCheckBox lanAllCheckBox;
    private javax.swing.JCheckBox lanBACheckBox;
    private javax.swing.JCheckBox lanBGCheckBox;
    private javax.swing.JCheckBox lanCZCheckBox;
    private javax.swing.JCheckBox lanDECheckBox;
    private javax.swing.JCheckBox lanDKCheckBox;
    private javax.swing.JCheckBox lanEECheckBox;
    private javax.swing.JCheckBox lanESCheckBox;
    private javax.swing.JCheckBox lanFICheckBox;
    private javax.swing.JCheckBox lanFRCheckBox;
    private javax.swing.JCheckBox lanGBCheckBox;
    private javax.swing.JCheckBox lanGRCheckBox;
    private javax.swing.JCheckBox lanHRCheckBox;
    private javax.swing.JCheckBox lanHUCheckBox;
    private javax.swing.JCheckBox lanISCheckBox;
    private javax.swing.JCheckBox lanITCheckBox;
    private javax.swing.JCheckBox lanLTCheckBox;
    private javax.swing.JCheckBox lanLVCheckBox;
    private javax.swing.JLabel lanLabel;
    private javax.swing.JCheckBox lanNLCheckBox;
    private javax.swing.JCheckBox lanPLCheckBox;
    private javax.swing.JCheckBox lanPTCheckBox;
    private javax.swing.JCheckBox lanROCheckBox;
    private javax.swing.JCheckBox lanRSCheckBox;
    private javax.swing.JCheckBox lanRUCheckBox;
    private javax.swing.JCheckBox lanSECheckBox;
    private javax.swing.JCheckBox lanSICheckBox;
    private javax.swing.JCheckBox lanSKCheckBox;
    private javax.swing.JCheckBox lanTRCheckBox;
    private javax.swing.JCheckBox lanUACheckBox;
    private javax.swing.JButton listBrowseButton;
    private javax.swing.JLabel listLabel;
    private javax.swing.JRadioButton listRadioButton;
    private javax.swing.JButton listStartButton;
    private javax.swing.JTextField listTextField;
    private javax.swing.JCheckBox oneFolderCheckBox;
    private javax.swing.JCheckBox orientLCheckBox;
    private javax.swing.JLabel orientLabel;
    private javax.swing.JCheckBox orientPCheckBox;
    private javax.swing.JCheckBox outDateCheckBox;
    private com.toedter.calendar.JDateChooser outDateChooser;
    private javax.swing.JLabel outExampleLabel;
    private javax.swing.JComboBox outExtComboBox;
    private javax.swing.JComboBox outItemComboBox;
    private javax.swing.JLabel outLabel;
    private javax.swing.JCheckBox outLangCheckBox;
    private javax.swing.JCheckBox outOrientCheckBox;
    private javax.swing.JCheckBox outPdfCheckBox;
    private javax.swing.JCheckBox outTextCheckBox;
    private javax.swing.JTextField outTextField;
    private javax.swing.JLabel shownEanLabel;
    private javax.swing.JRadioButton shownEanRadioButton;
    private javax.swing.JLabel shownItemLabel;
    private javax.swing.JRadioButton shownItemRadioButton;
    private javax.swing.JLabel shownLabel;
    private javax.swing.JLabel shownSapLabel;
    private javax.swing.JRadioButton shownSapRadioButton;
    private javax.swing.ButtonGroup shownbuttonGroup;
    private org.jdesktop.beansbinding.BindingGroup bindingGroup;
    // End of variables declaration//GEN-END:variables

}

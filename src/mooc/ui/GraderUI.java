/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.ui;

import java.awt.Color;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import mooc.grader.Grader;
import javax.swing.JOptionPane;

/**
 *
 * @author aavendan
 */
public class GraderUI extends javax.swing.JFrame {

    /**
     * Creates new form GraderUI
     */
    public GraderUI() {
        initComponents();
        fileName.setBackground(Color.white);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        lblTitle = new javax.swing.JLabel();
        lblOriginal = new javax.swing.JLabel();
        fileName = new javax.swing.JTextField();
        loadFile1 = new javax.swing.JButton();
        lblResponses = new javax.swing.JLabel();
        folderName = new javax.swing.JTextField();
        loadFile2 = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        txtResponses = new javax.swing.JTextArea();
        lblResponses1 = new javax.swing.JLabel();
        btnGrade = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        lblTitle.setFont(new java.awt.Font("Lucida Grande", 1, 24)); // NOI18N
        lblTitle.setText("MOOC Grader");

        lblOriginal.setFont(new java.awt.Font("Lucida Grande", 1, 13)); // NOI18N
        lblOriginal.setText("Archivo Original:");

        fileName.setDisabledTextColor(new java.awt.Color(0, 0, 0));
        fileName.setEnabled(false);

        loadFile1.setText("Cargar Archivo");
        loadFile1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cargarArchivo(evt);
            }
        });

        lblResponses.setFont(new java.awt.Font("Lucida Grande", 1, 13)); // NOI18N
        lblResponses.setText("Carpeta de Respuestas:");

        folderName.setDisabledTextColor(new java.awt.Color(0, 0, 0));
        folderName.setEnabled(false);

        loadFile2.setText("Cargar Archivo");
        loadFile2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cargarDirectorio(evt);
            }
        });

        txtResponses.setEditable(false);
        txtResponses.setColumns(20);
        txtResponses.setRows(5);
        jScrollPane1.setViewportView(txtResponses);

        lblResponses1.setFont(new java.awt.Font("Lucida Grande", 1, 13)); // NOI18N
        lblResponses1.setText("Archivos de Respuestas:");

        btnGrade.setText("Calificar");
        btnGrade.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                gradeResponses(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(12, 12, 12)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(lblResponses)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(lblOriginal)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(fileName)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(loadFile1, javax.swing.GroupLayout.PREFERRED_SIZE, 123, javax.swing.GroupLayout.PREFERRED_SIZE))))
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(6, 6, 6)
                                .addComponent(lblResponses1)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(folderName, javax.swing.GroupLayout.PREFERRED_SIZE, 516, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(loadFile2, javax.swing.GroupLayout.PREFERRED_SIZE, 123, javax.swing.GroupLayout.PREFERRED_SIZE))))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(217, 217, 217)
                                .addComponent(lblTitle))
                            .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 629, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(btnGrade)
                .addGap(270, 270, 270))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(16, 16, 16)
                .addComponent(lblTitle)
                .addGap(32, 32, 32)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblOriginal)
                    .addComponent(fileName, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(loadFile1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblResponses)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(folderName, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(loadFile2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblResponses1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 215, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnGrade)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void cargarArchivo(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cargarArchivo
        // TODO add your handling code here:
        if (evt.getSource() == loadFile1) {
            JFileChooser fc = new JFileChooser();
            fc.setCurrentDirectory(new java.io.File("/Users/aavendan/Documents/ESPOL/HCD/testGM"));
            fc.setDialogTitle("Seleccione el archivo original");
            int returnVal = fc.showOpenDialog(this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                //This is where a real application would open the file.
                fileName.setText(file.getAbsolutePath());
            }
        }
    }//GEN-LAST:event_cargarArchivo

    private void cargarDirectorio(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cargarDirectorio
        // TODO add your handling code here:
        if (evt.getSource() == loadFile2) {
            JFileChooser fc = new JFileChooser();
            fc.setCurrentDirectory(new java.io.File("/Users/aavendan/Documents/ESPOL/HCD/testGM"));
            fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            fc.setDialogTitle("Seleccione el directorio con respuestas");
            int returnVal = fc.showOpenDialog(this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                folderName.setText(fc.getSelectedFile().getAbsolutePath());
                listFiles(fc.getSelectedFile().getAbsolutePath());
            }
        }
    }//GEN-LAST:event_cargarDirectorio

    private void gradeResponses(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_gradeResponses
        // TODO add your handling code here:
        if (evt.getSource() == btnGrade && fileName.getText().length() > 0 && folderName.getText().length() > 0) {
            File fileFolder = new File(folderName.getText());
            File[] listOfFiles = fileFolder.listFiles();

            System.out.println(fileName.getText());
            for (int i = 0; i < listOfFiles.length; i++) {
                if (listOfFiles[i].isFile() && listOfFiles[i].getName().endsWith(".docx")) {
                    System.out.println(listOfFiles[i].getAbsolutePath());
                    Grader.printReport(fileName.getText(), listOfFiles[i].getAbsolutePath());
                }
            }
            
            JOptionPane.showMessageDialog(this, "¡Terminada!");
        }
    }//GEN-LAST:event_gradeResponses

    private void listFiles(String absolutePath) {
        File folder = new File(absolutePath);
        File[] listOfFiles = folder.listFiles();

        txtResponses.setText("");
        String filesNames = "";

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile() && listOfFiles[i].getName().endsWith(".docx")) {
                filesNames += listOfFiles[i].getName() + "\n";
            }
        }

        txtResponses.setText(filesNames);
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
            java.util.logging.Logger.getLogger(GraderUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(GraderUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(GraderUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(GraderUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new GraderUI().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnGrade;
    private javax.swing.JTextField fileName;
    private javax.swing.JTextField folderName;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JLabel lblOriginal;
    private javax.swing.JLabel lblResponses;
    private javax.swing.JLabel lblResponses1;
    private javax.swing.JLabel lblTitle;
    private javax.swing.JButton loadFile1;
    private javax.swing.JButton loadFile2;
    private javax.swing.JTextArea txtResponses;
    // End of variables declaration//GEN-END:variables

}

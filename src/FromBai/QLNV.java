/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package FromBai;

import java.io.FileOutputStream;
import java.io.ObjectOutputStream;
import Utils.NhanVien;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import Utils.ClockThread;
import Utils.XFile;
import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class QLNV extends javax.swing.JFrame {

    DefaultTableModel tblModel;
    List<NhanVien> list = new ArrayList<>();
    private int index = -1;
    private static final String P_EMAIL = "^([a-z0-9_\\.-]+)@([\\da-z\\.-]+)\\.([a-z\\.]{2,6})$";

    public QLNV() {
        initComponents();
        setLocationRelativeTo(null);
    }

    // fill tên cột
    public void TenCotNV() {
        tblModel = new DefaultTableModel();
        tblModel.addColumn("Mã");
        tblModel.addColumn("Họ và Tên");
        tblModel.addColumn("Tuổi");
        tblModel.addColumn("Email");
        tblModel.addColumn("Lương");
        tblNhanVien.setModel(tblModel);
    }

    public void fillTable() {
        tblModel.setRowCount(0);
        for (NhanVien nv : list) {
            tblModel.addRow(new Object[]{nv.getMa(), nv.getHoTen(), nv.getTuoi(), nv.getEmail(), nv.getLuong()});
        }
    }

//    public void Reset() {
//        txtHoTen.setText("");
//        txtEmail.setText("");
//        txtLuong.setText("");
//        txtMaNv.setText("");
//        txtTuoi.setText("");
//       // index = -1;
//    }

    public NhanVien readForm() {
        return new NhanVien(txtMaNv.getText(), txtHoTen.getText(), Integer.valueOf(txtTuoi.getText()), txtEmail.getText(), Double.parseDouble(txtLuong.getText()));
    }

    public void addNV() {
        if (validateForm()) {
            if (index == -1) {
                list.add(readForm());
                fillTable();
                JOptionPane.showMessageDialog(this, "Đã thêm");
            } else {
                capNhat(readForm());
                fillTable();
                JOptionPane.showMessageDialog(this, "Đã cập nhật");
            }
        }
    }

    public NhanVien timTheoMa(String ID) {
        for (NhanVien nv : list) {
            if (nv.getMa().equalsIgnoreCase(ID)) {
                return nv;
            }
        }
        return null;
    }

    public void capNhat(NhanVien newnv) {
        NhanVien nv1 = timTheoMa(newnv.getMa());
        if (nv1 != null) {
            nv1.setHoTen(newnv.getHoTen());
            nv1.setTuoi(newnv.getTuoi());
            nv1.setEmail(newnv.getEmail());
            nv1.setLuong(newnv.getLuong());
        }
    }
// Đồng hồ

    public String layThongTinBanGhi() {
        return "Record: " + (index + 1) + " of " + list.size();
    }

    public void HienThiNhanVien(int index) {
        txtMaNv.setText(list.get(index).getMa());
        txtHoTen.setText(list.get(index).getHoTen());
        txtEmail.setText(list.get(index).getEmail());
        txtTuoi.setText(String.valueOf(list.get(index).getTuoi()));
        txtLuong.setText(String.valueOf(list.get(index).getLuong()));
//        int viTri = tblNhanVien.getSelectedRow();
//        txtMaNv.setText(tblNhanVien.getValueAt(viTri, 0).toString());
//        txtHoTen.setText(tblNhanVien.getValueAt(viTri, 1).toString());
//        txtTuoi.setText(tblNhanVien.getValueAt(viTri, 2).toString());
//        txtEmail.setText(tblNhanVien.getValueAt(viTri, 3).toString());
//        txtLuong.setText(tblNhanVien.getValueAt(viTri, 4).toString());
    }

    public void fillNhanVienLenForm(NhanVien nv) {
        txtMaNv.setText(nv.getMa());
        txtHoTen.setText(nv.getHoTen());
        txtEmail.setText(nv.getEmail());
        txtTuoi.setText(String.valueOf(nv.getTuoi()));
        txtLuong.setText(String.valueOf(nv.getLuong()));
    }

    public void DeleteNV() {
        if (txtMaNv.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chọn Nhân Viên Cần Xóa");
            return;
        }
        int hoi = JOptionPane.showConfirmDialog(null, "Bạn có muốn xóa", "Xóa Nhân Viên", JOptionPane.YES_NO_OPTION);
        if (hoi == JOptionPane.YES_OPTION) {
            list.remove(index);
            fillTable();
            JOptionPane.showMessageDialog(this, "Đã Xóa");
            Reset();
        }
    }

    public void first() {
        if (list.size() != 0) {
            index = 0;
            updateInfo();
        }
    }

    public void last() {
        if (list.size() != 0) {
            index = list.size() - 1;
            updateInfo();
        }
    }

    public void pre() {
        if (index == 0) {
            last();
        } else {
            index--;
        }

        updateInfo();
    }

    public void next() {
        if (list.size() != 0) {
            if (index == list.size() - 1) {
                first();
            } else {
                index++;
            }

            updateInfo();
        }
    }

    // Xuất Excel
    public void exportExcel() {
        try {
            XSSFWorkbook fWorkbook = new XSSFWorkbook();
            XSSFSheet fSheet = fWorkbook.createSheet("danhSachNV");
            XSSFRow row = null;
            Cell cell = null;

            row = fSheet.createRow(0);

            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue("Mã");

            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue("Họ Và Tên");

            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue("Tuổi");

            cell = row.createCell(4, CellType.STRING);
            cell.setCellValue("Email");

            cell = row.createCell(5, CellType.STRING);
            cell.setCellValue("Lương");
            for (int i = 0; i < list.size(); i++) {

                row = fSheet.createRow(i + 1);

                cell = row.createCell(1, CellType.STRING);
                cell.setCellValue(list.get(i).getMa());

                cell = row.createCell(2, CellType.STRING);
                cell.setCellValue(list.get(i).getHoTen());

                cell = row.createCell(3, CellType.STRING);
                cell.setCellValue(list.get(i).getTuoi());

                cell = row.createCell(4, CellType.STRING);
                cell.setCellValue(list.get(i).getEmail());

                cell = row.createCell(5, CellType.STRING);
                cell.setCellValue(list.get(i).getLuong());
            }
            File file = new File(".//danhsachNV.xlsx");

            try {
                FileOutputStream fos = new FileOutputStream(file);
                fWorkbook.write(fos);
                fos.close();

            } catch (Exception e) {
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        JOptionPane.showMessageDialog(this, "Xuat thanh cong !");
    }

    private void updateInfo() {
        tblNhanVien.setRowSelectionInterval(index, index);
        HienThiNhanVien(index);
        lblBanGhi.setText(layThongTinBanGhi());
    }
//Bắt lỗi

    public boolean validateForm() {
        if (txtMaNv.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chưa nhập mã nhân viên");
            return false;
        }
        if (txtHoTen.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chưa nhập họ tên");
            return false;

        }
        if (txtTuoi.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chưa nhập tuổi");
            return false;
        }
        try {
            Integer.parseInt(txtTuoi.getText());
            if (Integer.parseInt(txtTuoi.getText()) < 16 || Integer.parseInt(txtTuoi.getText()) > 55) {
                JOptionPane.showMessageDialog(null, "Tuổi phải từ 16 đến 55 !");
                return false;
            }
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(this, "Tuổi phải là số");
            return false;
        }

        if (txtEmail.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chưa nhập email");
            return false;
        }
        Matcher matcher = Pattern.compile(P_EMAIL).matcher(txtEmail.getText());
        if (!matcher.matches()) {
            JOptionPane.showMessageDialog(this, "Email sai định dạng", "Lỗi", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        if (txtLuong.getText().equals("")) {
            JOptionPane.showMessageDialog(this, "Chưa nhập lương");
            return false;
        }
        try {
            Double.parseDouble(txtLuong.getText());
            if (Double.parseDouble(txtLuong.getText()) < 5000000) {
                JOptionPane.showMessageDialog(null, "Lương phải trên 5 triệu !");
                return false;
            }
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(this, "Lương phải là số", "Error", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        return true;
    }

    public void readFile() {
        try {
            list = (List<NhanVien>) XFile.readObj("list.data");
            fillTable();
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    public void saveFile() {
        try {
            XFile.writeObj("list.data", list);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel2 = new javax.swing.JLabel();
        txtMaNv = new javax.swing.JTextField();
        lbtTime = new javax.swing.JLabel();
        jPanel1 = new javax.swing.JPanel();
        btnReset = new javax.swing.JButton();
        btnLuu = new javax.swing.JButton();
        btnTim = new javax.swing.JButton();
        btnXoa = new javax.swing.JButton();
        btnThoat = new javax.swing.JButton();
        btnMo = new javax.swing.JButton();
        btnXuat = new javax.swing.JButton();
        jLabel4 = new javax.swing.JLabel();
        txtHoTen = new javax.swing.JTextField();
        jLabel5 = new javax.swing.JLabel();
        txtTuoi = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        txtEmail = new javax.swing.JTextField();
        jLabel7 = new javax.swing.JLabel();
        txtLuong = new javax.swing.JTextField();
        btnFirst = new javax.swing.JButton();
        btnPre = new javax.swing.JButton();
        btnNext = new javax.swing.JButton();
        btnLast = new javax.swing.JButton();
        lblBanGhi = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblNhanVien = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("QUẢN LÝ NHÂN VIÊN");
        setBackground(new java.awt.Color(153, 153, 255));
        setMinimumSize(new java.awt.Dimension(570, 472));
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel2.setText("MÃ NHÂN VIÊN");
        getContentPane().add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(12, 73, -1, -1));
        getContentPane().add(txtMaNv, new org.netbeans.lib.awtextra.AbsoluteConstraints(109, 70, 100, -1));

        lbtTime.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        lbtTime.setForeground(new java.awt.Color(255, 0, 51));
        lbtTime.setText("00:00 AM");
        getContentPane().add(lbtTime, new org.netbeans.lib.awtextra.AbsoluteConstraints(490, 10, -1, -1));

        jPanel1.setBackground(new java.awt.Color(255, 153, 0));
        jPanel1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(204, 204, 204)));

        btnReset.setText("New");
        btnReset.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnResetActionPerformed(evt);
            }
        });

        btnLuu.setText("Save");
        btnLuu.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLuuActionPerformed(evt);
            }
        });

        btnTim.setText("Find");
        btnTim.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnTimActionPerformed(evt);
            }
        });

        btnXoa.setText("Delete");
        btnXoa.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnXoaActionPerformed(evt);
            }
        });

        btnThoat.setText("Exit");
        btnThoat.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnThoatActionPerformed(evt);
            }
        });

        btnMo.setText("Open");
        btnMo.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnMoActionPerformed(evt);
            }
        });

        btnXuat.setText("Xuất Excel");
        btnXuat.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnXuatActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(btnXuat, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                    .addComponent(btnLuu, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnReset, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnThoat, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnMo, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnXoa, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 80, Short.MAX_VALUE)
                    .addComponent(btnTim, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(btnReset)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnLuu)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnXoa)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnTim)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnMo)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnThoat)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnXuat)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        getContentPane().add(jPanel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(464, 70, 106, -1));

        jLabel4.setText("HỌ VÀ TÊN");
        getContentPane().add(jLabel4, new org.netbeans.lib.awtextra.AbsoluteConstraints(31, 116, -1, -1));
        getContentPane().add(txtHoTen, new org.netbeans.lib.awtextra.AbsoluteConstraints(104, 113, 312, -1));

        jLabel5.setText("TUỔI");
        getContentPane().add(jLabel5, new org.netbeans.lib.awtextra.AbsoluteConstraints(66, 156, -1, -1));
        getContentPane().add(txtTuoi, new org.netbeans.lib.awtextra.AbsoluteConstraints(104, 153, 100, -1));

        jLabel6.setText("EMAIL");
        getContentPane().add(jLabel6, new org.netbeans.lib.awtextra.AbsoluteConstraints(58, 196, -1, -1));
        getContentPane().add(txtEmail, new org.netbeans.lib.awtextra.AbsoluteConstraints(104, 193, 312, -1));

        jLabel7.setText("LƯƠNG");
        getContentPane().add(jLabel7, new org.netbeans.lib.awtextra.AbsoluteConstraints(52, 236, -1, -1));
        getContentPane().add(txtLuong, new org.netbeans.lib.awtextra.AbsoluteConstraints(104, 233, 100, -1));

        btnFirst.setText("|<");
        btnFirst.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnFirstActionPerformed(evt);
            }
        });
        getContentPane().add(btnFirst, new org.netbeans.lib.awtextra.AbsoluteConstraints(92, 291, 41, -1));

        btnPre.setText("<<");
        btnPre.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnPreActionPerformed(evt);
            }
        });
        getContentPane().add(btnPre, new org.netbeans.lib.awtextra.AbsoluteConstraints(140, 291, 46, -1));

        btnNext.setText(">>");
        btnNext.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnNextActionPerformed(evt);
            }
        });
        getContentPane().add(btnNext, new org.netbeans.lib.awtextra.AbsoluteConstraints(193, 291, 46, -1));

        btnLast.setText(">|");
        btnLast.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLastActionPerformed(evt);
            }
        });
        getContentPane().add(btnLast, new org.netbeans.lib.awtextra.AbsoluteConstraints(246, 291, 41, -1));

        lblBanGhi.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        lblBanGhi.setForeground(new java.awt.Color(255, 0, 51));
        lblBanGhi.setText("Record: 1 of 10");
        getContentPane().add(lblBanGhi, new org.netbeans.lib.awtextra.AbsoluteConstraints(292, 295, -1, -1));

        jScrollPane1.setBackground(new java.awt.Color(255, 204, 102));

        tblNhanVien.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        tblNhanVien.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblNhanVienMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tblNhanVien);

        getContentPane().add(jScrollPane1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 330, 580, 176));

        jLabel1.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 0, 51));
        jLabel1.setText("QUẢN LÝ NHÂN VIÊN");
        getContentPane().add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(196, 13, -1, -1));

        jLabel3.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/bg_2.jpg"))); // NOI18N
        jLabel3.setText("jLabel3");
        getContentPane().add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 580, 520));

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnPreActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnPreActionPerformed
        pre();
    }//GEN-LAST:event_btnPreActionPerformed

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        // Thời gian
        ClockThread clockThread = new ClockThread(lbtTime);
        Thread t1 = new Thread(clockThread);
        t1.start();
//        list.add(new NhanVien("NV001", "Hiền", 18, "hieuntpc03470@fpt.edu.vn", 4000000));
//        list.add(new NhanVien("NV002", "Nam", 19, "namntpc03579@fpt.edu.vn", 2500000));
//        list.add(new NhanVien("NV003", "Quang", 20, "quangnmpc0000@fpt.edu.vn", 3300000));
//        list.add(new NhanVien("NV004", "Kỳ Anh", 35, "anhdnkpc08765@fpt.edu.vn", 1200000));
//        list.add(new NhanVien("NV005", "Nguyễn", 29, "nguyenhtpc08765@fpt.edu.vn", 1500000));

        TenCotNV();
        // fillTable();
        // lblBanGhi.setText(layThongTinBanGhi());
    }//GEN-LAST:event_formWindowOpened

    private void btnResetActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnResetActionPerformed
        Reset();
    }//GEN-LAST:event_btnResetActionPerformed

    private void btnLuuActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLuuActionPerformed
        addNV();
    }//GEN-LAST:event_btnLuuActionPerformed

    private void tblNhanVienMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblNhanVienMouseClicked
        index = tblNhanVien.getSelectedRow();
        HienThiNhanVien(index);
        lblBanGhi.setText(layThongTinBanGhi());
    }//GEN-LAST:event_tblNhanVienMouseClicked

    private void btnXoaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnXoaActionPerformed
        DeleteNV();
    }//GEN-LAST:event_btnXoaActionPerformed

    private void btnThoatActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnThoatActionPerformed
        try {
            int hoi = JOptionPane.showConfirmDialog(null, "Bạn có muốn Save", "Đăng xuất", JOptionPane.YES_NO_OPTION);
            if (hoi == JOptionPane.YES_OPTION) {
                FileOutputStream fos = new FileOutputStream("thoat");
                ObjectOutputStream oos = new ObjectOutputStream(fos);
                oos.writeObject(list);
                JOptionPane.showMessageDialog(null, "Save Thành Công");
                saveFile();
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Gặp Lỗi: " + e);
        }

        System.exit(0);
    }//GEN-LAST:event_btnThoatActionPerformed

    private void btnTimActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnTimActionPerformed
        String ma = JOptionPane.showInputDialog("mời bạn nhập mã");
        try {
            boolean kq = false;
            for (NhanVien x : list) {
                if (x.getMa().equalsIgnoreCase(ma)) {
                    index = list.indexOf(x);
                    HienThiNhanVien(index);
                    kq = true;
                    break;
                }

            }
            if (!kq) {
                JOptionPane.showMessageDialog(null, "Không tìm thấy nhân viên có mã " + ma);
            } else {
                JOptionPane.showMessageDialog(null, "Có nhân viên  " + ma);
            }

        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "gặp lỗi" + e);
        }
    }//GEN-LAST:event_btnTimActionPerformed

    private void btnFirstActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnFirstActionPerformed
        first();
    }//GEN-LAST:event_btnFirstActionPerformed

    private void btnLastActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLastActionPerformed
        last();
    }//GEN-LAST:event_btnLastActionPerformed

    private void btnNextActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnNextActionPerformed
        next();
    }//GEN-LAST:event_btnNextActionPerformed

    private void btnMoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnMoActionPerformed
        readFile();
    }//GEN-LAST:event_btnMoActionPerformed

    private void btnXuatActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnXuatActionPerformed
        exportExcel();
    }//GEN-LAST:event_btnXuatActionPerformed

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
            java.util.logging.Logger.getLogger(QLNV.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(QLNV.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(QLNV.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(QLNV.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new QLNV().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnFirst;
    private javax.swing.JButton btnLast;
    private javax.swing.JButton btnLuu;
    private javax.swing.JButton btnMo;
    private javax.swing.JButton btnNext;
    private javax.swing.JButton btnPre;
    private javax.swing.JButton btnReset;
    private javax.swing.JButton btnThoat;
    private javax.swing.JButton btnTim;
    private javax.swing.JButton btnXoa;
    private javax.swing.JButton btnXuat;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JLabel lblBanGhi;
    private javax.swing.JLabel lbtTime;
    private javax.swing.JTable tblNhanVien;
    private javax.swing.JTextField txtEmail;
    private javax.swing.JTextField txtHoTen;
    private javax.swing.JTextField txtLuong;
    private javax.swing.JTextField txtMaNv;
    private javax.swing.JTextField txtTuoi;
    // End of variables declaration//GEN-END:variables
}

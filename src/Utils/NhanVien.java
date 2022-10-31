
package Utils;

import java.io.Serializable;


public class NhanVien implements Serializable {
    private String ma, hoTen, email;
    private double luong;
    private int tuoi;

    public NhanVien() {
    }

    public NhanVien(String ma, String hoTen,int tuoi, String email, double luong ) {
        this.ma = ma;
        this.hoTen = hoTen;
        this.email = email;
        this.luong = luong;
        this.tuoi = tuoi;
    }

    public String getMa() {
        return ma;
    }

    public void setMa(String ma) {
        this.ma = ma;
    }

    public String getHoTen() {
        return hoTen;
    }

    public void setHoTen(String hoTen) {
        this.hoTen = hoTen;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public double getLuong() {
        return luong;
    }

    public void setLuong(double luong) {
        this.luong = luong;
    }

    public int getTuoi() {
        return tuoi;
    }

    public void setTuoi(int tuoi) {
        this.tuoi = tuoi;
    }
    
    
}

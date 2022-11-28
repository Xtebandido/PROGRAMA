package Modelo;

import java.util.Objects;

public class CLASE_anom1{
    private String anom1;

    public CLASE_anom1() {
        this.anom1 = this.anom1;
    }

    public String getAnom1() {
        return anom1;
    }
    public void setAnom1(String anom1) {
        this.anom1 = anom1;
    }

    @Override
    public String toString() {
        return "CLASE_anom1{" +
                "anom1='" + anom1 + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CLASE_anom1 that = (CLASE_anom1) o;
        return Objects.equals(anom1, that.anom1);
    }

    @Override
    public int hashCode() {
        return Objects.hash(anom1);
    }

}

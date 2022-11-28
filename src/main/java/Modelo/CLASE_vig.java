package Modelo;

import java.util.Objects;

public class CLASE_vig {
    private String vig;

    public CLASE_vig() {
        this.vig = this.vig;
    }

    public String getVig() {
        return vig;
    }
    public void setVig(String vig) {
        this.vig = vig;
    }

    @Override
    public String toString() {
        return "CLASE_vig{" +
                "vig='" + vig + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CLASE_vig that = (CLASE_vig) o;
        return Objects.equals(vig, that.vig);
    }

    @Override
    public int hashCode() {
        return Objects.hash(vig);
    }
}

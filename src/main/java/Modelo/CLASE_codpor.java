package Modelo;

import java.util.Objects;

public class CLASE_codpor {
    private String codpor;

    public CLASE_codpor() {
        this.codpor = this.codpor;
    }

    public String getCodpor() {
        return codpor;
    }
    public void setCodpor(String codpor) {
        this.codpor = codpor;
    }

    @Override
    public String toString() {
        return "CLASE_codpor{" +
                "codpor='" + codpor + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CLASE_codpor that = (CLASE_codpor) o;
        return Objects.equals(codpor, that.codpor);
    }

    @Override
    public int hashCode() {
        return Objects.hash(codpor);
    }
}

package Modelo;

import java.util.Objects;

public class CLASE_codope {
    private String codope;

    public CLASE_codope() {
        this.codope = this.codope;
    }

    public String getCodope() {
        return codope;
    }
    public void setCodope(String codope) {
        this.codope = codope;
    }

    @Override
    public String toString() {
        return "CLASE_codope{" +
                "codope='" + codope + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CLASE_codope that = (CLASE_codope) o;
        return Objects.equals(codope, that.codope);
    }

    @Override
    public int hashCode() {
        return Objects.hash(codope);
    }
}

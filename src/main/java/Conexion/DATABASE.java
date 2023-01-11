package Conexion;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import static Principal.PROGRAMA.pathDB;

public class DATABASE {

    public Connection conectarSQL() {
        Connection con = null;
        try {
            con = DriverManager.getConnection("jdbc:sqlite:" + pathDB.getAbsolutePath());
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    }

}

//HACER GRAFICAS DE LECTORES
//HACER GRAFICAS DE ANOMALIAS
//HACER UNI_LECTURA X LECTOR
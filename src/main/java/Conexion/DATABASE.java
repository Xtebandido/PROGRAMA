package Conexion;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DATABASE {
    File DB = new File("dbs\\BASE_DE_DATOS");

    public Connection conectarSQL() {
        Connection con = null;
        try {
            con = DriverManager.getConnection("jdbc:sqlite:" + DB.getAbsolutePath());
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    }
}

//HACER UN PANEL DONDE SELECCIONE LA BASE DE DATOS Y LA USE EN EL PROGRAMA
//HACER GRAFICAS DEL LECTOR
//
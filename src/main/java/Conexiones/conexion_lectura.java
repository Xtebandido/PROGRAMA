package Conexiones;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class conexion_lectura {
    File RutaDB = new File("dbs\\LECTURAS");
    String url = "" + RutaDB;

    public Connection conectarSQL() {
        Connection con = null;
        try {
            con = DriverManager.getConnection("jdbc:sqlite:" + url);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    }

}
package Principal; //PAQUETE PRINCIPAL
//CLASES Y LIBRERIAS IMPORTADAS
import Conexion.DATABASE;
import Modelo.*;
import com.aspose.cells.*;
import com.csvreader.CsvReader;
import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

//CLASE PRINCIPAL EXTENDIDA A JFRAME PARA LAS VISTAS IMPLEMENTANDO RUNNABLE PARA LAS TAREAS SINCRONICAS
public class PROGRAMA extends JFrame {
    //VARIABLES PRINCIPALES DE LA CLASE PROGRAMA
    JPanel mainPanel; //PANEL PRINCIPAL
    //------LOADING------
    JDialog dialog; //DIALOGO QUE CONTIENE LA PANTALLA DE CARGA
    //------INSERTAR-----
    JPanel jpIMPORT; //PANEL DE UNIR DENTRO DE PANEL DE LECTURAS
    //      ->SELECCIONAR
    JButton btnSELECT; //BOTON SELECCIONAR ARCHIVO
    File file = null; //DATO DONDE SE GUARDARA EL ARCHIVO SELECCIONADO
    JTextField jtxtPATH; //JTEXTFIELD CON EL DATO DE LA RUTA DEL ARCHIVO XLSX SELECCIONADO
    //      ->IMPORTAR
    String PATH = ""; //STRING QUE TIENE EL DATO DE LA RUTA DEL ARCHIVO SELECCIONADO PARA IMPORTAR
    JButton btnIMPORT; //BOTON IMPORTAR
    //--------EXPORTAR--------
    JPanel jpEXPORT; //PANEL DE EXPORTAR DENTRO DEL PANEL DE LECTURAS
    JButton btnEXPORT; // BOTON PARA EXPORTAR TODOS LOS DATOS
    int valINIT;
    int valFINISH;

    //METODO PRINCIPAL
    public PROGRAMA() {
        setContentPane(mainPanel);
        setTitle("ACUEDUCTO");
        setIconImage(new ImageIcon(getClass().getClassLoader().getResource("Multimedia/Icono.png")).getImage());
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        setVisible(true);

        //ACCION BOTON SELECCIONAR ARCHIVO
        btnSELECT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SELECTFILE();
            }
        });

        //ACCION BOTON IMPORTAR
        btnIMPORT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (PATH != "") {
                    new Thread(()-> FUN_IMPORT()).start();
                } else {
                    JOptionPane.showMessageDialog(null, "SELECCIONE UN ARCHIVO");
                }
            }
        });

        //ACCION BOTON EXPORTAR
        btnEXPORT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                new Thread(() -> CHECKING()).start();
                new Thread(() -> LOADING()).run();
            }
        });
    }

    //METODO LOADING
    public void LOADING() {
        JPanel panelLOAD; //PANEL DE CARGA
        JFrame frameLOAD; //FRAME DE CARGA
        JProgressBar pbCargar; //BARRA DE PROGRASO

        panelLOAD = new JPanel(new BorderLayout()); //PANEL DE CARGA
        frameLOAD = new JFrame(panelLOAD.getGraphicsConfiguration()); //NUEVO FRAME DE CARGA
        //BARRA DE PROGRESO INDEFINIDO
        pbCargar = new JProgressBar();
        pbCargar.setIndeterminate(true);
        //AÑADIR ELEMENTOS AL PANEL
        panelLOAD.add(new JLabel("CARGANDO REGISTROS... ESTO PODRIA TOMAR UNOS MINUTOS"), BorderLayout.PAGE_START); //AÑADIR UN LABEL AL INICIO DEL PANEL
        panelLOAD.add(pbCargar, BorderLayout.CENTER); //AÑADIR BARRA DE PROGRESO EN EL CENTRO DEL PANEL
        panelLOAD.setBackground(Color.CYAN); //ASIGNAR COLOR AZUL AL PANEL
        dialog = new JDialog(frameLOAD, true);

        dialog.setUndecorated(true);
        dialog.getContentPane().add(panelLOAD);
        dialog.pack();
        dialog.setLocationRelativeTo(null);
        dialog.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        dialog.setVisible(true);

    }

    //METODO SELECCIONAR ARCHIVO
    public void SELECTFILE() {
        JFileChooser fileChooser = new JFileChooser(file);
        if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            jtxtPATH.setText(fileChooser.getSelectedFile().toString());
            file = fileChooser.getCurrentDirectory(); // se guarda la ruta
            PATH = "" + fileChooser.getSelectedFile().toString();
        }
    }

    //METODO IMPORTAR
    public void FUN_IMPORT() {
    //1. CONVERTIR EL ARCHIVO.XLSX SELECCIONADO A ARCHIVO.CSV
        try {
            Workbook wbXLSX = new Workbook(PATH); //NUEVO LIBRO EXCEL
            Worksheet ws = wbXLSX.getWorksheets().get(0); //HOJA EXCEL, PRIMERA HOJA
            //VALIDAR ESTRUCTURA
            int valCOLUMN = ws.getCells().getMaxDataColumn(); //RECUENTO DE COLUMNA
            //SI TIENE 21 COLUMNAS HACER ESTO
            if ((valCOLUMN+1) == 21) {
                int valROW1 = ws.getCells().getLastDataRow(0); //RECUENTO DE COLUMNAS
                int valROW2 = ws.getCells().getMaxDataRow();
                if (valROW1 == valROW2) {
                    new Thread(() -> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
                    File fileDATA = new File("files\\Importe.csv"); //CREAR UN NUEVO ARCHIVO EN LA CARPETA files CON EL NOMBRE DE Importe DE TIPO csv
                    wbXLSX.save("" + fileDATA); //GUARDAR LOS DATOS DEL LIBRO EN EL ARCHIVO csv
                    String rutaCSV = "" + fileDATA; //GUARDAR RUTA EN UNA VARIABLE
                    // 2. LEE LOS DATOS DEL ARCHIVO Y LOS GUARDA EN UNA LISTA
                    List<LECTURAS> DATA; //LISTA CON MODELO DE LECTURAS LLAMADA DATA
                    DATA = new ArrayList<>(); //NUEVA LISTA DE DATOS DONDE SE GUARDARAN LOS DATOS DEL ARCHIVO

                    CsvReader readLECTURAS = new CsvReader(rutaCSV);
                    readLECTURAS.readHeaders();
                    //CICLO QUE LEE CADA DATO DEL ARCHIVO Y LOS ALMACENA EN LA LISTA
                    while (readLECTURAS.readRecord()) {
                        String codigo_porcion = readLECTURAS.get(0);
                        String uni_lectura = readLECTURAS.get(1);
                        String doc_lectura = readLECTURAS.get(2);
                        String cuenta_contrato = readLECTURAS.get(3);
                        String medidor = readLECTURAS.get(4);
                        String lectura_ant = readLECTURAS.get(5);
                        String lectura_act = readLECTURAS.get(6);
                        String anomalia_1 = readLECTURAS.get(7);
                        String anomalia_2 = readLECTURAS.get(8);
                        String codigo_operario = readLECTURAS.get(9);
                        String vigencia = readLECTURAS.get(10);
                        //CONVERTIR LOS DATOS RECIBIDOS DE fecha CON FORMATO yyyy/MM/dd HH:mm PARA MEJORAR LA FILTRACION
                        String fecha = readLECTURAS.get(11);
                        Calendar gregorianCalendar = new GregorianCalendar();
                        DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy HH:mm");
                        Date date = dateFormat.parse(fecha);
                        gregorianCalendar.setTime(date);
                        Locale locale = new Locale("es", "EC");
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm", locale);
                        fecha = simpleDateFormat.format(date);
                        //
                        String orden_lectura = readLECTURAS.get(12);
                        String leido = readLECTURAS.get(13);
                        String calle = readLECTURAS.get(14);
                        String edificio = readLECTURAS.get(15);
                        String suplemento_casa = readLECTURAS.get(16);
                        String interloc_comercial = readLECTURAS.get(17);
                        String apellido = readLECTURAS.get(18);
                        String nombre = readLECTURAS.get(19);
                        String clase_instalacion = readLECTURAS.get(20);

                        //SI EL DATO TIENE COMA, ELIMINARLA
                        codigo_porcion = codigo_porcion.replaceAll(",", "");
                        uni_lectura = uni_lectura.replaceAll(",", "");
                        doc_lectura = doc_lectura.replaceAll(",", "");
                        cuenta_contrato = cuenta_contrato.replaceAll(",", "");
                        medidor = medidor.replaceAll(",", "");
                        lectura_ant = lectura_ant.replaceAll(",", "");
                        lectura_act = lectura_act.replaceAll(",", "");
                        anomalia_1 = anomalia_1.replaceAll(",", "");
                        anomalia_2 = anomalia_2.replaceAll(",", "");
                        codigo_operario = codigo_operario.replaceAll(",", "");
                        vigencia = vigencia.replaceAll(",", "");
                        fecha = fecha.replaceAll(",", "");
                        orden_lectura = orden_lectura.replaceAll(",", "");
                        leido = leido.replaceAll(",", "");
                        calle = calle.replaceAll(",", "");
                        edificio = edificio.replaceAll(",", "");
                        suplemento_casa = suplemento_casa.replaceAll(",", "");
                        interloc_comercial = interloc_comercial.replaceAll(",", "");
                        apellido = apellido.replaceAll(",", "");
                        nombre = nombre.replaceAll(",", "");

                        //SI EL DATO TIENE COMILLAS, ELIMINARLAS
                        codigo_porcion = codigo_porcion.replaceAll("\"", "");
                        uni_lectura = uni_lectura.replaceAll("\"", "");
                        doc_lectura = doc_lectura.replaceAll("\"", "");
                        cuenta_contrato = cuenta_contrato.replaceAll("\"", "");
                        medidor = medidor.replaceAll("\"", "");
                        lectura_ant = lectura_ant.replaceAll("\"", "");
                        lectura_act = lectura_act.replaceAll("\"", "");
                        anomalia_1 = anomalia_1.replaceAll("\"", "");
                        anomalia_2 = anomalia_2.replaceAll("\"", "");
                        codigo_operario = codigo_operario.replaceAll("\"", "");
                        vigencia = vigencia.replaceAll("\"", "");
                        fecha = fecha.replaceAll("\"", "");
                        orden_lectura = orden_lectura.replaceAll("\"", "");
                        leido = leido.replaceAll("\"", "");
                        calle = calle.replaceAll("\"", "");
                        edificio = edificio.replaceAll("\"", "");
                        suplemento_casa = suplemento_casa.replaceAll("\"", "");
                        interloc_comercial = interloc_comercial.replaceAll("\"", "");
                        apellido = apellido.replaceAll("\"", "");
                        nombre = nombre.replaceAll("\"", "");

                        DATA.add(new LECTURAS(codigo_porcion, uni_lectura, doc_lectura, cuenta_contrato, medidor, lectura_ant, lectura_act, anomalia_1, anomalia_2, codigo_operario, vigencia, fecha, orden_lectura, leido, calle, edificio, suplemento_casa, interloc_comercial, apellido, nombre, clase_instalacion));
                    }
                    readLECTURAS.close();

                    //EXTRAER DATOS REPETIDOS DEL ARCHIVO
                    Set<LECTURAS> repetidos; //SET CON MODELO LECTURAS
                    repetidos = new HashSet<>(); //HASHSET PARA SACAR LOS REPETIDOS
                    List<LECTURAS> repetidosFinal; //LISTA CON MODELO LECTURAS
                    repetidosFinal = DATA.stream().filter(lectura -> !repetidos.add(lectura)).collect(Collectors.toList()); //GUARDAR DATOS REPETIDOS EN LA LISTA

                    boolean fileOPEN = false;
                    String name = jtxtPATH.getText();
                    name = name.replaceAll(" ", "_");
                    File fileNAME = new File(name);

                    //SI HAY REPETIDOS EXPORTARLOS EN UN EXCEL
                    if (repetidosFinal.size() != 0) {
                        File fileREPLY = new File("files\\Repetidos.csv"); //ARCHIVO PARA RETORNAR REPETIDOS EN UN ARCHIVO csv
                        PrintWriter write = new PrintWriter(fileREPLY); //PARA ESCRIBIR LOS DATOS REPETIDOS EN EL NUEVO ARCHIVO

                        String estructura = "CODIGO_PORCION,UNI_LECTURA,DOC_LECTURA,CUENTA_CONTRATO,MEDIDOR,LEC_ANTERIOR,LEC_ACTUAL,ANOMALIA_1,ANOMALIA_2,CODIGO_OPERARIO,VIGENCIA,FECHA,ORDEN LECTURA,LEIDO,CALLE,EDIFICIO,SUPLEMENTO_CASA,INTERLOC_COM,APELLIDO,NOMBRE,CLASE_INSTALA";
                        write.println(estructura);

                        for (Modelo.LECTURAS LECTURAS : repetidosFinal) {
                            write.print(LECTURAS.getCodigo_porcion() + ",");
                            write.print(LECTURAS.getUni_lectura() + ",");
                            write.print(LECTURAS.getDoc_lectura() + ",");
                            write.print(LECTURAS.getCuenta_contrato() + ",");
                            write.print(LECTURAS.getMedidor() + ",");
                            write.print(LECTURAS.getLectura_ant() + ",");
                            write.print(LECTURAS.getLectura_act() + ",");
                            write.print(LECTURAS.getAnomalia_1() + ",");
                            write.print(LECTURAS.getAnomalia_2() + ",");
                            write.print(LECTURAS.getCodigo_operario() + ",");
                            write.print(LECTURAS.getVigencia() + ",");
                            write.print(LECTURAS.getFecha() + ",");
                            write.print(LECTURAS.getOrden_lectura() + ",");
                            write.print(LECTURAS.getLeido() + ",");
                            write.print(LECTURAS.getCalle() + ",");
                            write.print(LECTURAS.getEdificio() + ",");
                            write.print(LECTURAS.getSuplemento_casa() + ",");
                            write.print(LECTURAS.getInterloc_comercial() + ",");
                            write.print(LECTURAS.getApellido() + ",");
                            write.print(LECTURAS.getNombre() + ",");
                            write.print(LECTURAS.getClase_instalacion());
                            write.println();
                        }
                        write.close();
                        //TRATAR DE CONVERTIR EL ARCHIVO.CSV CON DATOS REPETIDOS EN UN ARCHIVO.XLSX
                        try {
                            Workbook wbCSV = new Workbook("files\\Repetidos.csv"); //NUEVO LIBRO DEL ARCHIVO Repetidos
                            wbCSV.save("files\\REPETIDOS_" + fileNAME.getName(), SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
                        } catch (Exception e) {
                            fileOPEN = true;
                            dialog.dispose(); //CERRAR LOADING
                            JOptionPane.showMessageDialog(null, "ERROR: EL ARCHIVO NO PUEDE SER IMPORTADO PORQUE UN ARCHIVO RELACIONADO A LOS REGISTROS REPETIDOS SE ENCUENTRA ABIERTO", "", JOptionPane.INFORMATION_MESSAGE);
                        }
                        fileREPLY.delete(); //ELIMINAR ARCHIVO DE Repetidos.csv
                    }
                    fileDATA.delete(); //ELIMINAR ARCHIVO DE Importe.csv

                    if (fileOPEN != true) {
                        DATA = DATA.stream().distinct().collect(Collectors.toList()); //GUARDAR DATOS COMPLETOS SIN REPETIDOS
                        File RutaDATA = new File("files\\Datos.csv"); //ARCHIVO CON LOS DATOS COMPLETOS EN FORMATO csv
                        PrintWriter writeDATA = new PrintWriter(RutaDATA); //PARA ESCRIBIR LOS DATOS COMPLETOS EN EL NUEVO ARCHIVO

                        for (Modelo.LECTURAS LECTURAS : DATA) {
                            writeDATA.print(LECTURAS.getCodigo_porcion() + ",");
                            writeDATA.print(LECTURAS.getUni_lectura() + ",");
                            writeDATA.print(LECTURAS.getDoc_lectura() + ",");
                            writeDATA.print(LECTURAS.getCuenta_contrato() + ",");
                            writeDATA.print(LECTURAS.getMedidor() + ",");
                            writeDATA.print(LECTURAS.getLectura_ant() + ",");
                            writeDATA.print(LECTURAS.getLectura_act() + ",");
                            writeDATA.print(LECTURAS.getAnomalia_1() + ",");
                            writeDATA.print(LECTURAS.getAnomalia_2() + ",");
                            writeDATA.print(LECTURAS.getCodigo_operario() + ",");
                            writeDATA.print(LECTURAS.getVigencia() + ",");
                            writeDATA.print(LECTURAS.getFecha() + ",");
                            writeDATA.print(LECTURAS.getOrden_lectura() + ",");
                            writeDATA.print(LECTURAS.getLeido() + ",");
                            writeDATA.print(LECTURAS.getCalle() + ",");
                            writeDATA.print(LECTURAS.getEdificio() + ",");
                            writeDATA.print(LECTURAS.getSuplemento_casa() + ",");
                            writeDATA.print(LECTURAS.getInterloc_comercial() + ",");
                            writeDATA.print(LECTURAS.getApellido() + ",");
                            writeDATA.print(LECTURAS.getNombre() + ",");
                            writeDATA.print(LECTURAS.getClase_instalacion());
                            writeDATA.println();
                        }
                        writeDATA.close();

                        // 3. IMPORTAR LISTA DE DATOS A LA BASE DE DATOS
                        //CREAR ARCHIVO DE COMANDOS CON LAS RUTAS DE LA BASE DE DATOS Y EL ARCHIVO
                        File RutaDB = new File("dbs\\BASE_DE_DATOS");
                        File RutaCARPETA = new File("lib\\sqlite-tools");
                        File RutaCOMANDOS = new File("lib\\sqlite-tools\\comandos.txt");
                        PrintWriter writeCOMANDOS = new PrintWriter(RutaCOMANDOS); //PARA ESCRIBIR EL COMANDO CON LA RUTA DE LOS DATOS

                        //COMANDO (script)
                        writeCOMANDOS.println(".mode csv");
                        writeCOMANDOS.println(".open '" + RutaDB.getAbsolutePath() + "'");
                        writeCOMANDOS.println(".import '" + RutaDATA.getAbsolutePath() + "' LECTURAS");
                        writeCOMANDOS.println(".shell del '" + RutaDATA.getAbsolutePath() + "'");
                        writeCOMANDOS.close();

                        //LINEA DE COMANDOS EJECUTANDO EL COMANDO (script)
                        Runtime.getRuntime().exec("cmd /c cd " + RutaCARPETA.getAbsolutePath() + " && script.cmd");
                        Thread.sleep(2*1000);
                        dialog.dispose(); //CERRAR LOADING

                        JOptionPane.showMessageDialog(null, "SE IMPORTO CORRECTAMENTE " + DATA.size() + " REGISTROS DE " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                        if (repetidosFinal.size() != 0) {
                            JOptionPane.showMessageDialog(null, "SE ENCONTRARON " + repetidosFinal.size() + " REGISTROS REPETIDOS EN " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                            File rutaARCHIVOS = new File("files");
                            Runtime.getRuntime().exec("cmd /c start " + rutaARCHIVOS.getAbsolutePath() + "\\REPETIDOS_" + fileNAME.getName() + " && exit");
                        }
                        jtxtPATH.setText(null);
                        PATH = "";
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LOS DATOS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
                }
            } else {
                JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LA ESTRUCTURA DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR LA ESTRUCTURA DEL ARCHIVO
            }
        } catch (Exception e) {
            dialog.dispose(); //CERRAR LOADING
            File file = new File("files\\Importe.csv");
            file.delete();
            JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LOS DATOS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
        }
    }

    //METODO VALIDAR SI EL INFORME SE ENCUENTRA ABIERTO, VALIDAR LOS AÑOS DE LA BASE DE DATOS QUE SEAN UNICAMENTE LOS ULTIMOS 4 AÑOS E INICIAR LAS TAREAS PARA REALIZAR EL INFORME
    public void CHECKING(){
        boolean fileOPEN = false;
        try {
            Workbook wb = new Workbook(); //NUEVO LIBRO
            wb.save("files\\INFORME.xlsx"); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
        } catch (Exception e) {
            fileOPEN = true;
            dialog.dispose(); //CERRAR LOADING
            JOptionPane.showMessageDialog(null, "ERROR: EL INFORME NO PUEDE SER EXPORTADO PORQUE EL ARCHIVO SE ENCUENTRA ABIERTO. CIERRELO Y VUELVA A INTENTARLO", "", JOptionPane.INFORMATION_MESSAGE);
        }
        //SI EL ARCHIVO NO SE ENCUENTRA ABIERTO PROCEDER CON LA VERIFICACION
        if (fileOPEN != true) {
            //VALIDAR QUE LA DIFERENCIA DE VIGENCIA DE LECTURAS SEAN DE 4 AÑOS PARA EL INFORME
            DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
            Connection con = sql.conectarSQL(); //LLAMA LA CONEXION

            try {
                List<VIGENCIAS> Vigencias = new ArrayList<VIGENCIAS>();
                PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia");
                ResultSet rsVigencia = psVigencia.executeQuery();
                while (rsVigencia.next()) {
                    VIGENCIAS Vigencia = new VIGENCIAS();
                    Vigencia.setVigencia(rsVigencia.getString("vigencia"));
                    Vigencias.add(Vigencia);
                }
                int vigINICIAL;
                int vigFINAL = Integer.parseInt(Vigencias.get(Vigencias.size()-1).getVigencia());

                for (int j = 0; j < Vigencias.size(); j++) {
                    vigINICIAL = Integer.parseInt(Vigencias.get(j).getVigencia());
                    if ((vigFINAL - vigINICIAL) >= 400) {
                        Statement delete = con.createStatement();
                        delete.executeUpdate("DELETE FROM LECTURAS WHERE vigencia = '" + vigINICIAL + "'");
                    }
                }

                //INICIAR METODOS
                new Thread(() -> infoANOMALIAS()).start();
                new Thread(() -> infoCONSUMO_0()).start();
                new Thread(() -> infoLECTURAS()).start();

            } catch (Exception ex) {
            }
        }

    }

    //METODO informe -> ANOMALIAS
    public void infoANOMALIAS() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y FINALIZAR LA PANTALLA DE CARGA
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR VIGENCIAS
            List<VIGENCIAS> Vigencias = new ArrayList<VIGENCIAS>(); //LISTA CON MODELO DE VIGENCIAS
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia");
            ResultSet rsVigencia = psVigencia.executeQuery();
            while (rsVigencia.next()) {
                VIGENCIAS Vigencia = new VIGENCIAS();
                Vigencia.setVigencia(rsVigencia.getString("vigencia"));
                Vigencias.add(Vigencia);
            }

            //LISTAS ANOMALIAS
            List Anomalias = new ArrayList<Integer>();
            for (int i = 4; i <= 30; i++) {
                if (i == 22) {
                    i += 1;
                }
                Anomalias.add(i);
            }

            //LISTAR DESCRIPCION
            List Descripcion = new ArrayList<String>();
            Descripcion.add("MEDIDOR EN MAL ESTADO");
            Descripcion.add("MEDIDOR MAL INSTALADO");
            Descripcion.add("NÚMERO DE SERIE DE MEDIDOR NO CORRESPONDE");
            Descripcion.add("MEDIDOR SIN SELLOS O SELLOS ADULTERADOS");
            Descripcion.add("CAJILLA Y/O TAPA ROTA SUELTA Ó TRABADA");
            Descripcion.add("CAJILLA TAPADA O INUNDADA");
            Descripcion.add("SERVICIO DIRECTO");
            Descripcion.add("MEDIDOR RETIRADO");
            Descripcion.add("ACOMETIDA CON MEDIDOR Y POSIBLE CONEXIÓN FRAUDULENTA");
            Descripcion.add("ESCAPE EN LA ACOMETIDA");
            Descripcion.add("MEDIDOR INSTALADO POR DEBAJO DEL NIVEL NORMAL");
            Descripcion.add("CRUCE DE PLUMAS");
            Descripcion.add("MEDIDOR DENTRO DE PREDIO CERRADO CON LLAVE");
            Descripcion.add("NO SE LOCALIZA CAJILLA NI MEDIDOR");
            Descripcion.add("PREDIO DESOCUPADO");
            Descripcion.add("PREDIO NO LOCALIZADO EN TERRENO");
            Descripcion.add("PREDIO FUERA DE RUTA");
            Descripcion.add("DIRECCIÓN DESACTUALIZADA");
            Descripcion.add("CLASE DE USO DESACTUALIZADO");
            Descripcion.add("OBRA EN ACABADOS CON TPO");
            Descripcion.add("PREDIO DEMOLIDO Ó LOTE CON ACOMETIDA");
            Descripcion.add("SERVICIO SUSPENDIDO");
            Descripcion.add("SECTOR PELIGROSO");
            Descripcion.add("PREDIO OCUPADO");
            Descripcion.add("PREDIO MAL ENRUTADO");
            Descripcion.add("OCUPACION INDETERMINADA");

            //LISTAR ANOMALIAS X VIGENCIA
            List<ANOMXVIG> AnomaliasXVigencia = new ArrayList<ANOMXVIG>();
            List AXV = new ArrayList<String>();
            String AxV = "";
            int separar = 1;
            for (int i = 4; i <= 30; i++) {
                AnomaliasXVigencia.clear();
                int j = 0;
                if (i == 22) {
                    i += 1;
                }
                for (j = 0; j < Vigencias.size(); j++) {
                    PreparedStatement psANOMXVIG = con.prepareStatement("SELECT count(anomalia_1) as \"ANOMXVIG\" FROM LECTURAS WHERE ((anomalia_1 != \"\") AND (anomalia_1 =" + i + ") AND vigencia = '" + Vigencias.get(j).getVigencia() + "')");
                    ResultSet rsANOMXVIG = psANOMXVIG.executeQuery();
                    ANOMXVIG AnomXVig = new ANOMXVIG();
                    AnomXVig.setAnomXVig(rsANOMXVIG.getString("ANOMXVIG"));
                    AnomaliasXVigencia.add(AnomXVig);
                }

                for (ANOMXVIG model : AnomaliasXVigencia) {
                    AxV += model.getAnomXVig();
                    if (separar == Vigencias.size()) {
                        AXV.add(AxV);
                        separar = 1;
                    } else {
                        AxV += ",";
                        separar += 1;
                    }
                }
                AxV = "";
            }

            File fileANOMALIAS = new File("files\\ANOMALIAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeANOMALIAS = new PrintWriter(fileANOMALIAS); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "ANOMALIA,DESCRIPCION,";

            int separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += ("VIG" + Vigencia.getVigencia());
                if (0 < separadores) {
                    separadores--;
                    estructura += ",";
                }
            }
            writeANOMALIAS.println(estructura);

            for (int j = 0; j < Anomalias.size(); j++) {
                writeANOMALIAS.print(Anomalias.get(j) + ",");
                writeANOMALIAS.print(Descripcion.get(j));
                writeANOMALIAS.print("," + AXV.get(j));
                writeANOMALIAS.println();
            }

            writeANOMALIAS.close();

            //EXCEL
            //VARIABLES EXCEL
            Cell cell; //UNA CELDA
            Cells cells; //VARIAS CELDAS
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            Range range; //RANGO
            Border border; //BORDES

            //WORKBOOK
            Workbook wbANOMALIAS = new Workbook("files\\ANOMALIAS.csv"); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Worksheet wsANOMALIAS = wbANOMALIAS.getWorksheets().get(0); //NUEVA HOJA DE ANOMALIAS PARA EL LIBRO DE ANOMALIAS

            //ASIGNAR CELDAS CON UN TAMAÑO DEFINIDO
            cells = wsANOMALIAS.getCells();
            cells.setColumnWidth(0, 10);
            cells.setColumnWidth(1, 30);
            //ALINEAR CELDAS ANOMALIA Y DESCRIPCION A LA IZQUIERDA
            style = wbANOMALIAS.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.LEFT);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsANOMALIAS.getCells().createRange("A1:B28");
            range.applyStyle(style, flag);
            //ALINEAR CELDAS VIGENCIAS A LA DERECHA
            style = wbANOMALIAS.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.RIGHT);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsANOMALIAS.getCells().createRange("C1:Z28");
            range.applyStyle(style, flag);
            //COLOREAR CELDA B25 Y B26 DESCRIPCIONES DE LAS ANOMALIAS
            //B25
            cell = wsANOMALIAS.getCells().get("B25");
            style = cells.getStyle();
            style.setPattern(BackgroundType.SOLID);
            style.setForegroundColor(com.aspose.cells.Color.getYellow());
            cell.setStyle(style);
            //B26
            cell = wsANOMALIAS.getCells().get("B26");
            style = cell.getStyle();
            style.setPattern(BackgroundType.SOLID);
            style.setForegroundColor(com.aspose.cells.Color.getYellow());
            cell.setStyle(style);

            char c;
            int contador = 0;
            int columnas = 2;

            //FUNCION DE SUMAR LAS CELDAS DE CADA ANOMALIA X VIGENCIA EN FILA 28 SEGUN CADA COLUMNA DE VIGENCIA EXISTENTE
            for (c = 'C'; c <= 'Z'; ++c) {
                if (contador < Vigencias.size()) {
                    cells.setColumnWidth(columnas, 10.71);
                    cell = wsANOMALIAS.getCells().get(c + "28");
                    cell.setFormula("=SUM(" + c + "2:" + c + "27)");
                    columnas++;
                    contador++;
                }
            }

            //AGREGAR DISEÑO DE COLUMNAS COMO BORDES, TAMAÑO DE LETRA, TIPO DE LETRA Y COLORES

            contador = 0;
            columnas = 2;
            int fila = 1;
            columnas = columnas + Vigencias.size();

            for (c = 'A'; c <= 'Z'; ++c) {
                if (contador < columnas) {
                    for (fila = 1; fila <= 28; fila++) {
                        cell = wsANOMALIAS.getCells().get(c + "" + fila);
                        style = cell.getStyle();
                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                            cell.setStyle(style);
                        }

                        if (fila != 28) {
                            border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                            border.setLineStyle(CellBorderType.THIN);
                            cell.setStyle(style);
                            border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                            border.setLineStyle(CellBorderType.THIN);
                            cell.setStyle(style);
                            border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                            border.setLineStyle(CellBorderType.THIN);
                            cell.setStyle(style);
                            border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                            border.setLineStyle(CellBorderType.THIN);
                            cell.setStyle(style);
                        }
                        style.getFont().setName("Calibri");
                        cell.setStyle(style);
                        style.getFont().setSize(11);
                        cell.setStyle(style);

                    }
                    contador++;
                    fila = 1;
                }
            }
            con.close(); //CERRAR CONEXION
            wbANOMALIAS.save("files\\ANOMALIAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileANOMALIAS.delete(); //ELIMINAR ARCHIVO DE ANOMALIAS.csv

        } catch (Exception ex) {
        }
        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> CONSUMO_0
    public void infoCONSUMO_0() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y FINALIZAR LA PANTALLA DE CARGA
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR VIGENCIAS
            List<VIGENCIAS> Vigencias = new ArrayList<VIGENCIAS>();
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia");
            ResultSet rsVigencia = psVigencia.executeQuery();
            while (rsVigencia.next()) {
                VIGENCIAS Vigencia = new VIGENCIAS();
                Vigencia.setVigencia(rsVigencia.getString("vigencia"));
                Vigencias.add(Vigencia);
            }
            //LISTAR CODIGO PORCION & CONSUMO_0
            List<CON_0> Consumo_0 = new ArrayList<CON_0>();
            List Codigo_porcion = new ArrayList<String>();
            List CodPorXVig = new ArrayList<String>();
            String codporxvig = "";
            char c;
            int separar = 1;
            for (c = 'A'; c <= 'Z'; ++c) {
                Consumo_0.clear();
                String codpor = c + "4";
                if (codpor.equals("I4")) {
                    c = 'J';
                    codpor = "J4";
                } else if (codpor.equals("O4")) {
                    c = 'P';
                    codpor = "P4";
                } else if (codpor.equals("Y4")) {
                    c = 'Z';
                    codpor = "Z4";
                }
                Codigo_porcion.add(codpor);

                for (int i = 0; i < Vigencias.size(); i++) {
                    PreparedStatement psCON_0 = con.prepareStatement("SELECT count(*) AS CONSUMO_0 FROM LECTURAS WHERE (codigo_porcion = '" + codpor + "') AND (lectura_act - lectura_ant = 0) AND (lectura_act != '' AND lectura_ant != '') AND (vigencia = '" + Vigencias.get(i).getVigencia() + "')");
                    ResultSet rsCON_0 = psCON_0.executeQuery();
                    CON_0 con_0 = new CON_0();
                    con_0.setCon_0(rsCON_0.getString("CONSUMO_0"));
                    Consumo_0.add(con_0);
                }

                for (CON_0 model : Consumo_0) {
                    codporxvig += model.getCon_0();
                    if (separar == Vigencias.size()) {
                        CodPorXVig.add(codporxvig);
                        separar = 1;
                    } else {
                        codporxvig += ",";
                        separar += 1;
                    }
                }
                codporxvig = "";
            }

            File fileCONSUMO_0 = new File("files\\CONSUMO_0.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeCONSUMO_0 = new PrintWriter(fileCONSUMO_0); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "PORCION,";

            int separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += ("VIG" + Vigencia.getVigencia());
                if (0 < separadores) {
                    separadores--;
                    estructura += ",";
                }
            }
            writeCONSUMO_0.println(estructura);

            for (int j = 0; j < Codigo_porcion.size(); j++) {
                writeCONSUMO_0.print(Codigo_porcion.get(j));
                writeCONSUMO_0.print("," + CodPorXVig.get(j));
                writeCONSUMO_0.println();
            }
            writeCONSUMO_0.println("TOTAL");
            writeCONSUMO_0.close();

            //EXCEL
            //VARIABLES EXCEL
            Cell cell; //UNA CELDA
            Cells cells; //VARIAS CELDAS
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            Range range; //RANGO
            Border border; //BORDES

            //LIBRO EXCEL CONSUMO_0
            Workbook wbCONSUMO_0 = new Workbook("files\\CONSUMO_0.csv"); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Worksheet wsCONSUMO_0 = wbCONSUMO_0.getWorksheets().get(0); //NUEVA HOJA DE ANOMALIAS PARA EL LIBRO DE ANOMALIAS

            //ASIGNAR CELDAS CON UN TAMAÑO DEFINIDO
            cells = wsCONSUMO_0.getCells();
            cells.setColumnWidth(0, 11);
            //ALINEAR CELDAS PORCION A LA IZQUIERDA
            style = wbCONSUMO_0.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.LEFT);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsCONSUMO_0.getCells().createRange("A1:B25");
            range.applyStyle(style, flag);
            //ALINEAR CELDAS VIGENCIAS A LA DERECHA
            style = wbCONSUMO_0.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.CENTER);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsCONSUMO_0.getCells().createRange("B1:Z25");
            range.applyStyle(style, flag);

            int contador = 0;
            int columnas = 1;

            //FUNCION DE SUMAR LAS CELDAS DE CADA ANOMALIA X VIGENCIA EN FILA 28 SEGUN CADA COLUMNA DE VIGENCIA EXISTENTE
            for (c = 'B'; c <= 'Z'; ++c) {
                if (contador < Vigencias.size()) {
                    cells.setColumnWidth(columnas, 10);
                    cell = wsCONSUMO_0.getCells().get(c + "25");
                    cell.setFormula("=SUM(" + c + "2:" + c + "24)");
                    columnas++;
                    contador++;
                }
            }

            //AGREGAR DISEÑO DE COLUMNAS COMO BORDES, TAMAÑO DE LETRA, TIPO DE LETRA Y COLORES
            contador = 0;
            columnas = 1;
            columnas = columnas + Vigencias.size();

            for (c = 'A'; c <= 'Z'; ++c) {
                if (contador < columnas) {
                    for (int fila = 1; fila <= 25; fila++) {
                        cell = wsCONSUMO_0.getCells().get("A" + fila);
                        style = cell.getStyle();
                        style.setPattern(BackgroundType.SOLID);
                        style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                        cell.setStyle(style);
                    }
                    for (int fila = 1; fila <= 25; fila++) {
                        cell = wsCONSUMO_0.getCells().get(c + "" + fila);
                        style = cell.getStyle();
                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                            cell.setStyle(style);
                        }


                        border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);

                        style.getFont().setName("Calibri");
                        cell.setStyle(style);
                        style.getFont().setSize(11);
                        cell.setStyle(style);

                    }
                    contador++;
                }
            }
            con.close(); //CERRAR CONEXION
            wbCONSUMO_0.save("files\\CONSUMO_0.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileCONSUMO_0.delete(); //ELIMINAR ARCHIVO DE CONSUMO_0.csv
        } catch (Exception ex) {
        }
        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> LECTURAS
    public void infoLECTURAS() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y FINALIZAR LA PANTALLA DE CARGA
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR VIGENCIAS
            List<VIGENCIAS> Vigencias = new ArrayList<VIGENCIAS>();
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia");
            ResultSet rsVigencia = psVigencia.executeQuery();
            while (rsVigencia.next()) {
                VIGENCIAS Vigencia = new VIGENCIAS();
                Vigencia.setVigencia(rsVigencia.getString("vigencia"));
                Vigencias.add(Vigencia);
            }

            //LISTAR CODIGO PORCION & LEIDO, NO LEIDO, TOTAL Y LECTURAS
            List Codigo_porcion = new ArrayList<String>();
            List<LEIDO> Leido = new ArrayList<LEIDO>();
            List<NO_LEIDO> NoLeido = new ArrayList<NO_LEIDO>();
            List<TOTAL> Total = new ArrayList<TOTAL>();
            List Lecturas = new ArrayList<String>();
            String lnt = "";
            char c;
            int separar = 1;

            for(c = 'A'; c <= 'Z'; ++c) {
                Leido.clear();
                NoLeido.clear();
                Total.clear();

                String codpor = c + "4";
                if (codpor.equals("I4")) {
                    c = 'J';
                    codpor = "J4";
                }
                else if (codpor.equals("O4")) {
                    c = 'P';
                    codpor = "P4";
                } else if (codpor.equals("Y4")) {
                    c = 'Z';
                    codpor = "Z4";
                }
                Codigo_porcion.add(codpor);

                for (int i = 0; i < Vigencias.size(); i++) {

                    //LEIDO
                    PreparedStatement psLEIDO = con.prepareStatement("SELECT count(*) as LEIDO FROM LECTURAS WHERE (codigo_porcion = '" + codpor + "') AND lectura_act != '' AND vigencia = '" + Vigencias.get(i).getVigencia() + "'");
                    ResultSet rsLEIDO = psLEIDO.executeQuery();
                    LEIDO leido = new LEIDO();
                    leido.setLeido(rsLEIDO.getString("LEIDO"));
                    Leido.add(leido);

                    //NO LEIDO
                    PreparedStatement psNO_LEIDO = con.prepareStatement("SELECT count(*) as NO_LEIDO FROM LECTURAS WHERE (codigo_porcion = '" + codpor + "') AND lectura_act = '' AND vigencia = '" + Vigencias.get(i).getVigencia() + "'");
                    ResultSet rsNO_LEIDO = psNO_LEIDO.executeQuery();
                    NO_LEIDO no_leido = new NO_LEIDO();
                    no_leido.setNo_Leido(rsNO_LEIDO.getString("NO_LEIDO"));
                    NoLeido.add(no_leido);

                    //TOTAL
                    PreparedStatement psTOTAL = con.prepareStatement("SELECT count(*) as TOTAL FROM LECTURAS WHERE (codigo_porcion = '" + codpor + "') AND vigencia = '" + Vigencias.get(i).getVigencia() + "'");
                    ResultSet rsTOTAL = psTOTAL.executeQuery();
                    TOTAL total = new TOTAL();
                    total.setTotal(rsTOTAL.getString("TOTAL"));
                    Total.add(total);
                }

                for (int j = 0; j < Vigencias.size(); j++) {
                    lnt += Leido.get(j).getLeido() + "," + NoLeido.get(j).getNo_Leido() + "," + Total.get(j).getTotal();
                    if (separar == Vigencias.size()) {
                        Lecturas.add(lnt);
                        separar = 1;
                    } else {
                        lnt += ",";
                        separar += 1;
                    }
                }
                lnt = "";
            }

            File fileLECTURAS = new File("files\\LECTURAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeLECTURAS = new PrintWriter(fileLECTURAS); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = ",";

            int separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += ("VIG "+Vigencia.getVigencia());
                if (0 < separadores) {
                    separadores--;
                    estructura += ",,,";
                }
            }
            writeLECTURAS.println(estructura);

            estructura = "PORCION,";

            separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += "LEIDO,NO LEIDO,TOTAL ";
                if (0 < separadores) {
                    separadores--;
                    estructura += ",";
                }
            }
            writeLECTURAS.println(estructura);

            for (int j = 0; j < Codigo_porcion.size(); j++) {
                writeLECTURAS.print(Codigo_porcion.get(j));
                writeLECTURAS.print("," + Lecturas.get(j));
                writeLECTURAS.println();
            }
            writeLECTURAS.println("Total general");
            writeLECTURAS.close();

            //EXCEL
            //VARIABLES EXCEL
            Cell cell; //UNA CELDA
            Cells cells; //VARIAS CELDAS
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            Range range; //RANGO
            Border border; //BORDES

            Workbook wbLECTURAS = new Workbook("files\\LECTURAS.csv"); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Worksheet wsLECTURAS = wbLECTURAS.getWorksheets().get(0); //NUEVA HOJA DE ANOMALIAS PARA EL LIBRO DE ANOMALIAS

            //ASIGNAR CELDAS CON UN TAMAÑO DEFINIDO
            cells = wsLECTURAS.getCells();
            cells.setColumnWidth(0, 12);
            //ALINEAR CELDAS PORCION A LA IZQUIERDA
            style = wbLECTURAS.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.LEFT);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsLECTURAS.getCells().createRange("A2:B26");
            range.applyStyle(style, flag);
            //ALINEAR CELDAS VIGENCIAS A LA DERECHA
            style = wbLECTURAS.createStyle();
            style.setHorizontalAlignment(TextAlignmentType.CENTER);
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            flag.setAlignments(true);
            range = wsLECTURAS.getCells().createRange("B1:BU26");
            range.applyStyle(style, flag);

            //FUNCION DE SUMAR LAS CELDAS DE CADA ANOMALIA X VIGENCIA EN FILA 28 SEGUN CADA COLUMNA DE VIGENCIA EXISTENTE
            int contador = 0;
            int columnas = 1;
            int tamañoXvigencia = 0;

            while (contador < Vigencias.size()) {
                tamañoXvigencia += 3;
                contador++;
            }

            contador = 0;
            //SUMAR RANGO DE CELDAS DE LAS COLUMNAS AA HASTA BU
            for (c = 'A'; c <= 'Z'; ++c) {
                //COLOREAR DE AZUL LA COLUMNA DE CODIGO PORCION
                for (int fila = 2; fila <= 26; fila++) {
                    cell = wsLECTURAS.getCells().get("A" + fila);
                    style = cell.getStyle();
                    style.setPattern(BackgroundType.SOLID);
                    style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                    cell.setStyle(style);
                    //AGREGAR BORDES
                    border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                    border.setLineStyle(CellBorderType.THIN);
                    cell.setStyle(style);
                    border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                    border.setLineStyle(CellBorderType.THIN);
                    cell.setStyle(style);
                    border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                    border.setLineStyle(CellBorderType.THIN);
                    cell.setStyle(style);
                    border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                    border.setLineStyle(CellBorderType.THIN);
                    cell.setStyle(style);
                    //CAMBIAR TIPO DE FUENTE
                    style.getFont().setName("Calibri");
                    cell.setStyle(style);
                    style.getFont().setSize(11);
                    cell.setStyle(style);
                }

                if (contador == 0) {
                    c = 'B';
                }

                if (contador < tamañoXvigencia && contador <= 25) {
                    cells.setColumnWidth(columnas, 9.50);
                    cell = wsLECTURAS.getCells().get(c+"26");
                    cell.setFormula("=SUM("+c+"3:"+c+"25)");
                    columnas++;
                    contador++;

                    for (int fila = 1; fila <= 26; fila++) {
                        cell = wsLECTURAS.getCells().get(c + "" + fila);
                        style = cell.getStyle();

                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142));
                            cell.setStyle(style);
                        }
                        if (fila == 2) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                            cell.setStyle(style);
                        }

                        border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);

                        style.getFont().setName("Calibri");
                        cell.setStyle(style);
                        style.getFont().setSize(11);
                        cell.setStyle(style);
                    }
                    if (contador == 25) {
                        c = 'A';
                    }
                    //COMBINAR Y CENTRAR LA PRIMERA VIGENCIA
                    cells.merge(0, 1, 1, 3);
                    //COMBINAR Y CENTRAR 3 COLUMNAS POR CADA VIGENCIA
                    if ((contador%3) == 0 && contador < tamañoXvigencia) {
                        cells.merge(0, (contador+1), 1, 3);
                    }

                }

                if (contador >= 25 && contador < tamañoXvigencia && contador <= 51) {
                    cells.setColumnWidth(columnas, 9.50);
                    cell = wsLECTURAS.getCells().get("A" + c + "26");
                    cell.setFormula("=SUM(A" + c + "3:A" + c + "25)");
                    columnas++;
                    contador++;

                    for (int fila = 1; fila <= 26; fila++) {
                        cell = wsLECTURAS.getCells().get("A" + c + "" + fila);
                        style = cell.getStyle();

                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142));
                            cell.setStyle(style);
                        }
                        if (fila == 2) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                            cell.setStyle(style);
                        }

                        border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);

                        style.getFont().setName("Calibri");
                        cell.setStyle(style);
                        style.getFont().setSize(11);
                        cell.setStyle(style);
                    }

                    if (contador == 51) {
                        c = 'A';
                    }

                    //COMBINAR Y CENTRAR 3 COLUMNAS POR CADA VIGENCIA
                    if ((contador%3) == 0 && contador < tamañoXvigencia) {
                        cells.merge(0, (contador+1), 1, 3);
                    }

                }
                if (contador >= 51 && contador < tamañoXvigencia) {
                    cells.setColumnWidth(columnas, 9.50);
                    cell = wsLECTURAS.getCells().get("B" + c + "26");
                    cell.setFormula("=SUM(B" + c + "3:B" + c + "25)");
                    columnas++;
                    contador++;

                    for (int fila = 1; fila <= 26; fila++) {
                        cell = wsLECTURAS.getCells().get("B" + c + "" + fila);
                        style = cell.getStyle();

                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142));
                            cell.setStyle(style);
                        }
                        if (fila == 2) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219));
                            cell.setStyle(style);
                        }

                        border = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.LEFT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.RIGHT_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);
                        border = style.getBorders().getByBorderType(BorderType.TOP_BORDER);
                        border.setLineStyle(CellBorderType.THIN);
                        cell.setStyle(style);

                        style.getFont().setName("Calibri");
                        cell.setStyle(style);
                        style.getFont().setSize(11);
                        cell.setStyle(style);

                    }

                    String limit;
                    limit = "" + c;
                    if (limit.equals("U")) {
                        c = 'Z';
                    }

                    //COMBINAR Y CENTRAR 3 COLUMNAS POR CADA VIGENCIA
                    if ((contador%3) == 0 && contador < tamañoXvigencia) {
                        cells.merge(0, (contador+1), 1, 3);
                    }

                }

            }
            con.close(); //CERRAR CONEXION

            wbLECTURAS.save("files\\LECTURAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileLECTURAS.delete(); //ELIMINAR ARCHIVO DE CONSUMO_0.csv

        } catch (Exception ex) {
        }
        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO GENERAR INFORME
    public void INFORME() {

        try {
            //CREAR EXCEL DE INFORME
            Workbook wbINFORME = new Workbook(); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            //SELECCIONAR LOS LIBROS CON LAS TABLAS
            File workbook1 = new File("files\\ANOMALIAS.xlsx");
            File workbook2 = new File("files\\CONSUMO_0.xlsx");
            File workbook3 = new File("files\\LECTURAS.xlsx");
            Workbook wbCONSUMO_0 = new Workbook(workbook1.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook wbLECTURAS = new Workbook(workbook2.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook wbANOMALIAS = new Workbook(workbook3.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS

            //COMBINAR HOJAS EN EL INFORME
            wbINFORME.combine(wbCONSUMO_0);
            wbINFORME.combine(wbLECTURAS);
            wbINFORME.combine(wbANOMALIAS);
            wbINFORME.getWorksheets().removeAt(0); //TEMPORAL MIENTRAS ACABA TODAS LAS FUNCIONES PARA EL INFORME
            wbINFORME.save("files\\INFORME.xlsx");
            //ELIMINAR LIBROS COPIADOS
            workbook1.delete();
            workbook2.delete();
            workbook3.delete();

            dialog.dispose(); //CERRAR LOADING
            JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE EL INFORME");
            File ARCHIVOS = new File("files");
            Runtime.getRuntime().exec("cmd /c start " + ARCHIVOS.getAbsolutePath() + " && exit");
        } catch (Exception ex) {
        }

    }

    //METODO MAIN
    public static void main(String[] args) {
        new PROGRAMA();
    }

}


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

//CLASE PRINCIPAL (main) EXTENDIDA A JFRAME PARA LAS VISTAS
public class PROGRAMA extends JFrame implements Runnable {

    //VARIABLES PRINCIPALES DE LA CLASE PROGRAMA
    JPanel mainPanel; //PANEL PRINCIPAL

    //LOADING
    JPanel pCargar; //PANEL DE CARGA
    JFrame fCargar; //FRAME DE CARGA
    JProgressBar pbCargar; //BARRA DE PROGRESO

    //LECTURAS
    JPanel tpLecturas; //PANEL DE LECTURAS
    int btnOPRIMIDO; //ENTERO QUE RETORNA UN VALOR PARA IDENTIFICAR EL BOTON OPRIMIDO MEDIANTE UN SWITCH

    //------INSERTAR-----
    JPanel jpUnir; //PANEL DE UNIR DENTRO DE PANEL DE LECTURAS
    //->SELECCIONAR
    JButton btnSELECT; //BOTON SELECCIONAR ARCHIVO
    JTextField jtxtPATH; //JTEXTFIELD CON EL DATO DE LA RUTA DEL ARCHIVO XLSX SELECCIONADO
    //->IMPORTAR
    String PATH = ""; //STRING QUE TIENE EL DATO DE LA RUTA DEL ARCHIVO SELECCIONADO PARA IMPORTAR
    JButton btnIMPORT; //BOTON IMPORTAR
    File ARCHIVOS = new File("files");

    //--------EXPORTAR--------
    JPanel jpExportar; //PANEL DE EXPORTAR DENTRO DEL PANEL DE LECTURAS
    JButton btnEXPORT; // BOTON PARA EXPORTAR TODOS LOS DATOS

    List<VIGENCIAS> Vigencias; //LISTA CON MODELO DE ANOMALIAS LLAMADA Anomalias
    List<ANOMXVIG> AnomaliasXVigencia; //LISTA CON MODELO DE ANOMALIAS LLAMADA Anomalias
    List<CON_0> Consumo_0; //LISTA CON MODELO DE ANOMALIAS LLAMADA Anomalias


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
                    btnOPRIMIDO = 1;
                    run();
                } else {
                    JOptionPane.showMessageDialog(null, "SELECCIONE UN ARCHIVO");
                }
            }
        });

        //ACCION BOTON EXPORTAR TOD0
        btnEXPORT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                btnOPRIMIDO = 2;
                run();
            }
        });
    }

    //METODOS

    //METODO LOADING
    public void LOADING() {
        pbCargar = new JProgressBar(); //NUEVA BARRA DE PROGRESO
        pbCargar.setIndeterminate(true); //BARRA DE PROGRESO INDETERMINADA
        pCargar = new JPanel(new BorderLayout()); //NUEVO PANEL DE CARGA
        fCargar = new JFrame(pCargar.getGraphicsConfiguration()); //NUEVO FRAME DE CARGA
        //PANEL
        pCargar.add(new JLabel("CARGANDO REGISTROS. POR FAVOR, ESPERE...\n"), BorderLayout.CENTER);
        pCargar.add(pbCargar, BorderLayout.AFTER_LAST_LINE);
        pCargar.setBackground(Color.CYAN);
        //FRAME
        fCargar.setUndecorated(true);
        fCargar.getContentPane().add(pCargar);
        fCargar.pack();
        fCargar.setLocationRelativeTo(null);
        fCargar.setAlwaysOnTop(true);
        fCargar.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        fCargar.setVisible(true);
    }
    //SELECCIONAR ARCHIVO
    public void SELECTFILE() {
        File file = null; //NUEVO ARCHIVO DONDE SE GUARDARA EL ARCHIVO QUE SEA SELECCIONADO
        JFileChooser fileChooser = new JFileChooser(); //JFILECHOOSER FRAME DONDE SE SELECCIONA UN ARCHIVO
        fileChooser.showOpenDialog(null); //ABRE DIALOGO PARA SELECCIONAR ARCHIVO
        file = fileChooser.getSelectedFile(); //GUARDA EL ARCHIVO SELECCIONADO EN LA VARIABLE archivoSeleccionado
        //SI EL ARCHIVO FUE SELECCIONADO HACER ESTO
        if (file != null) {
            jtxtPATH.setText("" + file); //MOSTRAR LA RUTA DEL ARCHIVO EN EL JTEXTFIELD
            PATH = "" + file;
        }
    }

    //METODO run IMPLEMENTANDO EL RUNNABLE PARA INICIAR THREADS/MULTITAREA
    public void run() {
        switch (btnOPRIMIDO) {
            case 1:
                new Thread(()-> FUN_IMPORT()).start();
                break;
            case 2:
                new Thread(()-> EXPORT()).start();
                break;
        }
    }

    //METODO IMPORTAR
    public void FUN_IMPORT() {
    //1. CONVERTIR EL ARCHIVO.XLSX SELECCIONADO A ARCHIVO.CSV
        try {
            Workbook wbXLSX = new Workbook(PATH); //NUEVO LIBRO EXCEL
            Worksheet worksheet = wbXLSX.getWorksheets().get(0); //HOJA EXCEL, PRIMERA HOJA
            //VALIDAR ESTRUCTURA
            int cCols = 0; //INICIALIZAR VARIABLE CANTIDAD DE COLUMNAS
            int FILA = 1;
            String DATOCONCOMA = "";
            boolean valDATA = false; //BOOLEANO QUE VALIDA SI LOS DATOS CONTIENEN COMA (TRUE) O NO (FALSE)
            cCols = worksheet.getCells().getMaxDataColumn(); //RECUENTO DE COLUMNAS
            cCols = cCols + 1; //COMO INICIALIZA EN 0 ENTONCES SE SUMA 1 PARA QUE SE ACOMODE LA CANTIDAD DE COLUMNAS REQUERIDAS
            //SI TIENE 21 COLUMNAS HACER ESTO
            if (cCols == 21) {
                new Thread(()-> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
                File file = new File("files\\Importe.csv"); //CREAR UN NUEVO ARCHIVO EN LA CARPETA files CON EL NOMBRE DE Importe DE TIPO csv
                wbXLSX.save("" + file); //GUARDAR LOS DATOS DEL LIBRO EN EL ARCHIVO csv
                String rutaCSV = "" + file; //GUARDAR RUTA EN UNA VARIABLE
    // 2. LEE LOS DATOS DEL ARCHIVO Y LOS GUARDA EN UNA LISTA
                List<LECTURAS> DATA; //LISTA CON MODELO DE LECTURAS LLAMADA DATA
                DATA = new ArrayList<>(); //NUEVA LISTA DE DATOS DONDE SE GUARDARAN LOS DATOS DEL ARCHIVO
                try {
                    CsvReader readLECTURAS = new CsvReader(rutaCSV);
                    readLECTURAS.readHeaders();
                    //CICLO QUE LEE CADA DATO DEL ARCHIVO Y LOS ALMACENA EN LA LISTA
                    while (readLECTURAS.readRecord()) {
                        FILA = FILA + 1;
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
                        if (codigo_porcion.contains(",") || uni_lectura.contains(",") || doc_lectura.contains(",") || cuenta_contrato.contains(",") || medidor.contains(",") || lectura_ant.contains(",") || lectura_act.contains(",") || anomalia_1.contains(",") || anomalia_2.contains(",") || codigo_operario.contains(",") || vigencia.contains(",") || fecha.contains(",") || orden_lectura.contains(",") || leido.contains(",") || calle.contains(",") || edificio.contains(",") || suplemento_casa.contains(",") || interloc_comercial.contains(",") || apellido.contains(",") || nombre.contains(",") || clase_instalacion.contains(",")) {

                            calle = calle.replaceAll(",","");
                            edificio = edificio.replaceAll(",","");
                            suplemento_casa = suplemento_casa.replaceAll(",","");
                            interloc_comercial = interloc_comercial.replaceAll(",", "");
                            apellido = apellido.replaceAll(",", "");
                            nombre = nombre.replaceAll(",", "");

                            //VALIDAR NUEVAMENTE SI ALGUN OTRO DATO TIENE COMA Y DEVOLVER ERROR
                            if (codigo_porcion.contains(",") || uni_lectura.contains(",") || doc_lectura.contains(",") || cuenta_contrato.contains(",") || medidor.contains(",") || lectura_ant.contains(",") || lectura_act.contains(",") || anomalia_1.contains(",") || anomalia_2.contains(",") || codigo_operario.contains(",") || vigencia.contains(",") || fecha.contains(",") || orden_lectura.contains(",") || leido.contains(",") || calle.contains(",") || edificio.contains(",") || suplemento_casa.contains(",") || interloc_comercial.contains(",") || apellido.contains(",") || nombre.contains(",") || clase_instalacion.contains(",")) {
                                fCargar.dispose(); //CERRAR LOADING
                                valDATA = true;
                                DATOCONCOMA += "FILA " + FILA + " → " + codigo_porcion + " | " + uni_lectura + " | " + doc_lectura + " | " + cuenta_contrato + " | " + medidor + " | " + lectura_ant + " | " + lectura_act + " | " + anomalia_1 + " | " + anomalia_2 + " | " + codigo_operario + " | " + vigencia + " | " + fecha + " | " + orden_lectura + " | " + leido + " | " + calle + " | " + edificio + " | " + suplemento_casa + " | " + interloc_comercial + " | " + apellido + " | " + nombre + " | " + clase_instalacion + "\n";

                            }
                        }
                        DATA.add(new LECTURAS(codigo_porcion, uni_lectura, doc_lectura, cuenta_contrato, medidor, lectura_ant, lectura_act, anomalia_1, anomalia_2, codigo_operario, vigencia, fecha, orden_lectura, leido, calle, edificio, suplemento_casa, interloc_comercial, apellido, nombre, clase_instalacion));
                    }
                    readLECTURAS.close();

                    if (valDATA == true){
                        file.delete();
                    }
                    if (valDATA == false) {
                        //EXTRAER DATOS REPETIDOS DEL ARCHIVO
                        Set<LECTURAS> repetidos; //SET CON MODELO LECTURAS LLAMADA repetidos
                        repetidos = new HashSet<>(); //HASHSET
                        List<LECTURAS> repetidosFinal; //LISTA CON MODELO LECTURAS LLAMADA repetidosFinal

                        repetidosFinal = DATA.stream().filter(lectura -> !repetidos.add(lectura)).collect(Collectors.toList());

                        File csvFile = new File("files\\Repetidos.csv"); //ARCHIVO PARA RETORNAR REPETIDOS EN UN ARCHIVO csv
                        PrintWriter write = new PrintWriter(csvFile); //PARA ESCRIBIR LOS DATOS REPETIDOS EN EL NUEVO ARCHIVO

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
                        //CONVERTIR EL ARCHIVO.CSV CON DATOS REPETIDOS EN UN ARCHIVO.XLSX
                        Workbook wbCSV = new Workbook("files\\Repetidos.csv"); //NUEVO LIBRO DEL ARCHIVO Repetidos
                        wbCSV.save("files\\Repetidos.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
                        file.delete(); //ELIMINAR ARCHIVO DE Importe
                        csvFile.delete(); //ELIMINAR ARCHIVO DE Repetidos

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
                        writeCOMANDOS.print(".shell del '" + RutaDATA.getAbsolutePath() + "'");
                        writeCOMANDOS.close();
                        //LINEA DE COMANDOS EJECUTANDO EL COMANDO (script)
                        Runtime.getRuntime().exec("cmd /c start cmd.exe /K \" cd " + RutaCARPETA.getAbsolutePath() + " && script.cmd && exit");

                        fCargar.dispose(); //CERRAR LOADING
                        if (repetidosFinal.size() == 0) {
                            File RutaREPETIDOS = new File("files\\Repetidos.xlsx");
                            RutaREPETIDOS.delete();
                            JOptionPane.showMessageDialog(null, "NO SE ENCONTRO NINGUN REGISTRO REPETIDO EN EL ARCHIVO");
                        } else {
                            JOptionPane.showMessageDialog(null, "SE ENCONTRO " + repetidosFinal.size() + " REGISTROS REPETIDOS EN EL ARCHIVO");
                            Runtime.getRuntime().exec("cmd /c start cmd.exe /K \" start " + ARCHIVOS.getAbsolutePath() + "\\Repetidos.xlsx" + " && exit");
                        }
                        jtxtPATH.setText(null);
                        PATH = "";
                        JOptionPane.showMessageDialog(null, "SE IMPORTO CORRECTAMENTE " + DATA.size() + " REGISTROS");
                    }
                    else {
                        JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LOS SIGUIENTES DATOS DEL ARCHIVO:\n"+DATOCONCOMA); //MENSAJE DE ERROR POR LA ESTRUCTURA DEL ARCHIVO
                    }
                } catch(Exception e) {
                    throw new RuntimeException(e);
                }
            } else {
                JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LA ESTRUCTURA DEL ARCHIVO"); //MENSAJE DE ERROR POR LA ESTRUCTURA DEL ARCHIVO
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    //3. METODOS PARA EXPORTAR
    public void EXPORT(){
        new Thread(()-> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
        new Thread(()-> INFORME()).start(); //INICIAR TAREA ANOMALIAS
    }

    public void INFORME() {
        //VARIABLES EXCEL
        Cell cell; //UNA CELDA
        Cells cells; //VARIAS CELDAS
        Style style; //ESTILO
        StyleFlag flag = new StyleFlag(); //BANDERA
        Range range; //RANGO
        Border border; //BORDES

        //ANOMALIAS
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION

        //LISTAS NECESARIAS PARA OBTENER ANOMALIAS X VIGENCIAS EXISTENTES
        List Anomalias = new ArrayList<Integer>();
        List Descripcion = new ArrayList<String>();
        Vigencias = new ArrayList<VIGENCIAS>();
        AnomaliasXVigencia = new ArrayList<ANOMXVIG>();
        List AXV = new ArrayList<String>();
        String AxV = "";

        try {
            //AGREGAR ANOMALIAS
            for (int i = 4; i <= 30; i++) {
                if (i == 22) {
                    i += 1;
                }
                Anomalias.add(i);
            }
            //AGREGAR DESCRIPCION
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

            //AGREGAR VIGENCIAS
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia");
            ResultSet rsVigencia = psVigencia.executeQuery();
            while (rsVigencia.next()) {
                VIGENCIAS Vigencia = new VIGENCIAS();
                Vigencia.setVigencia(rsVigencia.getString("vigencia"));
                Vigencias.add(Vigencia);
            }

            //CONTEO ANOMALIAS X VIGENCIA
            int separar = 1;
            for (int i = 4; i <= 30; i++) {
                AnomaliasXVigencia.clear();
                int j = 0;
                if (i == 22){
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
                estructura += ("VIG"+Vigencia.getVigencia());
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

            //EXCEL ANOMALIAS
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
            for(c = 'C'; c <= 'Z'; ++c) {
                if (contador < Vigencias.size()) {
                    cells.setColumnWidth(columnas, 10.71);
                    cell = wsANOMALIAS.getCells().get(c+"28");
                    cell.setFormula("=SUM("+c+"2:"+c+"27)");
                    columnas++;
                    contador++;
                }
            }

            contador = 0;
            columnas = 2;
            int fila = 1;
            columnas = columnas + Vigencias.size();

            //AGREGAR DISEÑO DE COLUMNAS COMO BORDES, TAMAÑO DE LETRA, TIPO DE LETRA Y COLORES
            for(c = 'A'; c <= 'Z'; ++c) {
                if (contador < columnas) {
                    for (fila = 1; fila <= 28; fila++) {
                        cell = wsANOMALIAS.getCells().get(c+""+fila);
                        style = cell.getStyle();
                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142,169,219));
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

            wbANOMALIAS.save("files\\ANOMALIAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileANOMALIAS.delete(); //ELIMINAR ARCHIVO DE ANOMALIAS.csv
            //FIN ANOMALIAS

            //CONSUMO 0
            List Codigo_porcion = new ArrayList<String>();
            Consumo_0 = new ArrayList<CON_0>();
            String codpor;

            for(c = 'A'; c <= 'Z'; ++c) {
                codpor = c + "4";
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
            }
            for (int i = 0; i < Vigencias.size(); i++){
                for (int j = 0; j < Codigo_porcion.size(); j++) {
                    PreparedStatement psCON_0 = con.prepareStatement("SELECT count(*) AS CONSUMO_0 FROM LECTURAS WHERE codigo_porcion = '"+Codigo_porcion.get(j)+"' AND lectura_act - lectura_ant = 0 AND vigencia = '" + Vigencias.get(i).getVigencia() + "'");
                    ResultSet rsCON_0 = psCON_0.executeQuery();
                    CON_0 con_0 = new CON_0();
                    con_0.setCon_0(rsCON_0.getString("CONSUMO_0"));
                    Consumo_0.add(con_0);
                }
            }

            File fileCONSUMO_0 = new File("files\\CONSUMO_0.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeCONSUMO_0 = new PrintWriter(fileCONSUMO_0); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            estructura = "PORCION,";

            separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += ("VIG"+Vigencia.getVigencia());
                if (0 < separadores) {
                    separadores--;
                    estructura += ",";
                }
            }
            writeCONSUMO_0.println(estructura);

            for (int j = 0; j < Codigo_porcion.size(); j++) {
                writeCONSUMO_0.print(Codigo_porcion.get(j));
                writeCONSUMO_0.print("," + Consumo_0.get(j).getCon_0());
                writeCONSUMO_0.println();
            }
            writeCONSUMO_0.println("TOTAL");
            writeCONSUMO_0.close();

            //EXCEL CONSUMO_0
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

            contador = 0;
            columnas = 1;

            //FUNCION DE SUMAR LAS CELDAS DE CADA ANOMALIA X VIGENCIA EN FILA 28 SEGUN CADA COLUMNA DE VIGENCIA EXISTENTE
            for(c = 'B'; c <= 'Z'; ++c) {
                if (contador < Vigencias.size()) {
                    cells.setColumnWidth(columnas, 10);
                    cell = wsCONSUMO_0.getCells().get(c+"25");
                    cell.setFormula("=SUM("+c+"2:"+c+"24)");
                    columnas++;
                    contador++;
                }
            }

            contador = 0;
            columnas = 1;
            fila = 1;
            columnas = columnas + Vigencias.size();

            //AGREGAR DISEÑO DE COLUMNAS COMO BORDES, TAMAÑO DE LETRA, TIPO DE LETRA Y COLORES
            for(c = 'A'; c <= 'Z'; ++c) {
                if (contador < columnas) {
                    for (fila = 1; fila <= 25; fila++) {
                        cell = wsCONSUMO_0.getCells().get("A"+fila);
                        style = cell.getStyle();
                        style.setPattern(BackgroundType.SOLID);
                        style.setForegroundColor(com.aspose.cells.Color.fromArgb(142,169,219));
                        cell.setStyle(style);
                    }
                    for (fila = 1; fila <= 25; fila++) {
                        cell = wsCONSUMO_0.getCells().get(c+""+fila);
                        style = cell.getStyle();
                        if (fila == 1) {
                            style.setPattern(BackgroundType.SOLID);
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142,169,219));
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
                    fila = 1;
                }
            }

            wbCONSUMO_0.save("files\\CONSUMO_0.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileCONSUMO_0.delete(); //ELIMINAR ARCHIVO DE CONSUMO_0.csv
            //FIN CONSUMO_0


            con.close();
            fCargar.dispose(); // CERRAR PANTALLA DE CARGA
            JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE EL INFORME");
            Runtime.getRuntime().exec("cmd /c start cmd.exe /K \" start " + ARCHIVOS.getAbsolutePath() + " && exit");

        } catch (Exception ex) {
        }
    }

    //METODO MAIN
    public static void main(String[] args) {
        new PROGRAMA();
    }
}


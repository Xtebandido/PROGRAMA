package Principal; //PAQUETE PRINCIPAL
//CLASES Y LIBRERIAS IMPORTADAS
import Conexiones.*;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import com.toedter.calendar.JDateChooser;

//CLASE PRINCIPAL (main) EXTENDIDA A JFRAME PARA LAS VISTAS
public class PROGRAMA extends JFrame implements Runnable {

    //VARIABLES PRINCIPALES DE LA CLASE PROGRAMA
    JPanel mainPanel; //PANEL PRINCIPAL
    JTabbedPane mainTabbedPanne; //TABBEDPANE DONDE ESTAN LOS PANELES EN PESTAÑAS

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

    //---------FILTRAR---------
    JPanel jpFiltrar; //PANEL DE FILTRAR DENTRO DEL PANEL DE LECTURAS
    JButton btnFCodPorcion; //BOTON PARA FILTRAR CODIGO PORCION
    JButton btnFAnomalia1; //BOTON PARA FILTRAR ANOMALIA 1
    JButton btnFCodOperario; //BOTON PARA FILTRAR CODIGO OPERARIO
    JButton btnFVigencia; //BOTON PARA FILTRAR VIGENCIA
    JButton btnFFecha; //BOTON PARA FILTRAR FECHA

    //VARIABLES PARA FILTRAR CODIGO PORCION
    List<CLASE_codpor> listCodPor; //LISTA CON MODELO DE CLASE_codpor LLAMADA listCodPor
    JPanel panelCODPOR = new JPanel(new BorderLayout()); //PANEL CON EL CONTENIDO QUE SE VERA EN EL FRAME
    JFrame frameCODPOR = new JFrame((panelCODPOR.getGraphicsConfiguration())); //FRAME QUE CONTIENE LA VISTA FILTRAR DE CODIGO PORCION
    String queryFilCodPorcion = ""; //STRING QUE GUARDA EL DATO DE LOS DATOS FILTRADOS DE LOS CAMPOS CODIGO PORCION
    int contCodPor; //ENTERO QUE FUNCIONA PARA VALIDAR SI ESTA ABIERTO O CERRADO EL PANEL DEL CAMPO

    //VARIABLES PARA FILTRAR ANOMALIA 1
    List<CLASE_anom1> listANOM1; //LISTA CON MODELO DE CLASE_anom1 LLAMADA listANOM1
    JPanel panelANOM1 = new JPanel(new BorderLayout()); //PANEL PARA FILTRAR ANOMALIA 1
    JFrame frameANOM1 = new JFrame((panelANOM1.getGraphicsConfiguration())); //FRAME PARA FILTRAR ANOMALIA 1
    String queryFilAnomalia1 = ""; //STRING QUE GUARDA EL DATO DE LOS DATOS FILTRADOS DE LOS CAMPOS ANOMALIA 1
    int contANOM1; //ENTERO QUE FUNCIONA PARA VALIDAR SI ESTA ABIERTO O CERRADO EL PANEL DEL CAMPO

    //VARIABLES PARA FILTRAR CODIGO OPERARIO
    List<CLASE_codope> listCodOpe; //LISTA CON MODELO DE CLASE_codope LLAMADA listCodOpe
    JPanel panelCODOPE = new JPanel(new BorderLayout()); //PANEL PARA FILTRAR CODIGO OPERARIO
    JFrame frameCODOPE = new JFrame((panelCODOPE.getGraphicsConfiguration())); //FRAME PARA FILTRAR CODIGO OPERARIO
    String queryFilCodOperario = ""; //STRING QUE GUARDA EL DATO DE LOS DATOS FILTRADOS DE LOS CAMPOS CODIGO OPERARIO
    int contCodOpe; //ENTERO QUE FUNCIONA PARA VALIDAR SI ESTA ABIERTO O CERRADO EL PANEL DEL CAMPO

    //VARIABLES PARA FILTRAR VIGENCIA
    List<CLASE_vig> listVig; //LISTA CON MODELO DE CLASE_vig LLAMADA listVig
    JPanel panelVIG = new JPanel(new BorderLayout()); //PANEL PARA FILTRAR VIGENCIA
    JFrame frameVIG = new JFrame((panelVIG.getGraphicsConfiguration())); //FRAME PARA FILTRAR VIGENCIA
    String queryFilVig = ""; //STRING QUE GUARDA EL DATO DE LOS DATOS FILTRADOS DE LOS CAMPOS VIGENCIA
    int contVig; //ENTERO QUE FUNCIONA PARA VALIDAR SI ESTA ABIERTO O CERRADO EL PANEL DEL CAMPO

    //VARIABLES PARA FILTRAR FECHA
    JPanel panelFEC = new JPanel(new BorderLayout()); //PANEL PARA FILTRAR FECHA
    JFrame frameFEC = new JFrame((panelFEC.getGraphicsConfiguration())); //FRAME PARA FILTRAR FECHA
    String rangoDesde; //STRING QUE GUARDA EL DATO FILTRADO DE UNA FECHA INICIAL
    String rangoHasta; //STRING QUE GUARDA EL DATO FILTRADO DE UNA FECHA FINAL
    String queryFilFec = ""; //STRING QUE GUARDA EL DATO DE LOS DATOS FILTRADOS DE LOS CAMPOS FECHA
    int contFec; //ENTERO QUE FUNCIONA PARA VALIDAR SI ESTA ABIERTO O CERRADO EL PANEL DEL CAMPO

    //VARIABLES PARA VALIDAR LOS FRAMES DE FILTRACIONES
    int contValCodPor;
    int contValANOM1;
    int contValCodOpe;
    int contValVig;
    int contValFec;

    //--------EXPORTAR--------
    JPanel jpExportar; //PANEL DE EXPORTAR DENTRO DEL PANEL DE LECTURAS
    JButton btnExportarFil; //BOTON PARA EXPORTAR DATOS FILTRADOS
    JButton btnExportarAll; // BOTON PARA EXPORTAR TODOS LOS DATOS

    List<VIGENCIAS> Vigencias; //LISTA CON MODELO DE ANOMALIAS LLAMADA Anomalias
    List<ANOMXVIG> AnomaliasXVigencia; //LISTA CON MODELO DE ANOMALIAS LLAMADA Anomalias

    List<getLECTURAS> DATOSdb; //LISTA CON MODELO DE getLECTURAS LLAMADA DATOSdb

    //METODO PRINCIPAL
    public PROGRAMA() {
        setContentPane(mainPanel);
        setTitle("ACUEDUCTO");
        setIconImage(new ImageIcon(getClass().getClassLoader().getResource("Multimedia/Icono.png")).getImage());
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        setVisible(true);

        //INICIALIZAR CONTADORES EN 1 PARA LOS FRAMES DE FILTRAR
        contCodPor = 1;
        contANOM1 = 1;
        contCodOpe = 1;
        contVig = 1;
        contFec = 1;

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

        //ACCION BOTON CODIGO PORCION
        btnFCodPorcion.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                contCodPor = contCodPor + 1;
                if (contCodPor % 2 == 0) {
                    filCODPOR();
                } else {
                    frameCODPOR.dispose();
                }

            }
        });
        //ACCION BOTON ANOMALIA 1
        btnFAnomalia1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                contANOM1 = contANOM1 + 1;
                if (contANOM1 % 2 == 0) {
                    filANOM1();
                } else {
                    frameANOM1.dispose();
                }
            }
        });
        //ACCION BOTON CODIGO OPERARIO
        btnFCodOperario.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                contCodOpe = contCodOpe + 1;
                if (contCodOpe % 2 == 0) {
                    filCODOPE();
                } else {
                    frameCODOPE.dispose();
                }
            }
        });
        //ACCION BOTON VIGENCIA
        btnFVigencia.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                contVig = contVig + 1;
                if (contVig % 2 == 0) {
                    filVIG();
                } else {
                    frameVIG.dispose();
                }
            }
        });
        //ACCION BOTON FECHA
        btnFFecha.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                contFec = contFec + 1;
                if (contFec % 2 == 0) {
                    filFEC();
                } else {
                    frameFEC.dispose();
                }
            }
        });
        //ACCION BOTON EXPORTAR TOD0
        btnExportarAll.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                btnOPRIMIDO = 2;
                run();
            }
        });
        //ACCION BOTON EXPORTAR FILTRADOS
        btnExportarFil.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                btnOPRIMIDO = 3;
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
            case 3:
                new Thread(()-> EXPORTARfiltrados()).start();
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
                List<setLECTURAS> DATA; //LISTA CON MODELO DE LECTURAS LLAMADA DATA
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
                        DATA.add(new setLECTURAS(codigo_porcion, uni_lectura, doc_lectura, cuenta_contrato, medidor, lectura_ant, lectura_act, anomalia_1, anomalia_2, codigo_operario, vigencia, fecha, orden_lectura, leido, calle, edificio, suplemento_casa, interloc_comercial, apellido, nombre, clase_instalacion));
                    }
                    readLECTURAS.close();

                    if (valDATA == true){
                        file.delete();
                    }
                    if (valDATA == false) {
                        //EXTRAER DATOS REPETIDOS DEL ARCHIVO
                        Set<setLECTURAS> repetidos; //SET CON MODELO LECTURAS LLAMADA repetidos
                        repetidos = new HashSet<>(); //HASHSET
                        List<setLECTURAS> repetidosFinal; //LISTA CON MODELO LECTURAS LLAMADA repetidosFinal

                        repetidosFinal = DATA.stream().filter(lectura -> !repetidos.add(lectura)).collect(Collectors.toList());

                        File csvFile = new File("files\\Repetidos.csv"); //ARCHIVO PARA RETORNAR REPETIDOS EN UN ARCHIVO csv
                        PrintWriter write = new PrintWriter(csvFile); //PARA ESCRIBIR LOS DATOS REPETIDOS EN EL NUEVO ARCHIVO

                        String estructura = "CODIGO_PORCION,UNI_LECTURA,DOC_LECTURA,CUENTA_CONTRATO,MEDIDOR,LEC_ANTERIOR,LEC_ACTUAL,ANOMALIA_1,ANOMALIA_2,CODIGO_OPERARIO,VIGENCIA,FECHA,ORDEN LECTURA,LEIDO,CALLE,EDIFICIO,SUPLEMENTO_CASA,INTERLOC_COM,APELLIDO,NOMBRE,CLASE_INSTALA";
                        write.println(estructura);

                        for (setLECTURAS setLECTURAS : repetidosFinal) {
                            write.print(setLECTURAS.getCodigo_porcion() + ",");
                            write.print(setLECTURAS.getUni_lectura() + ",");
                            write.print(setLECTURAS.getDoc_lectura() + ",");
                            write.print(setLECTURAS.getCuenta_contrato() + ",");
                            write.print(setLECTURAS.getMedidor() + ",");
                            write.print(setLECTURAS.getLectura_ant() + ",");
                            write.print(setLECTURAS.getLectura_act() + ",");
                            write.print(setLECTURAS.getAnomalia_1() + ",");
                            write.print(setLECTURAS.getAnomalia_2() + ",");
                            write.print(setLECTURAS.getCodigo_operario() + ",");
                            write.print(setLECTURAS.getVigencia() + ",");
                            write.print(setLECTURAS.getFecha() + ",");
                            write.print(setLECTURAS.getOrden_lectura() + ",");
                            write.print(setLECTURAS.getLeido() + ",");
                            write.print(setLECTURAS.getCalle() + ",");
                            write.print(setLECTURAS.getEdificio() + ",");
                            write.print(setLECTURAS.getSuplemento_casa() + ",");
                            write.print(setLECTURAS.getInterloc_comercial() + ",");
                            write.print(setLECTURAS.getApellido() + ",");
                            write.print(setLECTURAS.getNombre() + ",");
                            write.print(setLECTURAS.getClase_instalacion());
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

                        for (setLECTURAS setLECTURAS : DATA) {
                            writeDATA.print(setLECTURAS.getCodigo_porcion() + ",");
                            writeDATA.print(setLECTURAS.getUni_lectura() + ",");
                            writeDATA.print(setLECTURAS.getDoc_lectura() + ",");
                            writeDATA.print(setLECTURAS.getCuenta_contrato() + ",");
                            writeDATA.print(setLECTURAS.getMedidor() + ",");
                            writeDATA.print(setLECTURAS.getLectura_ant() + ",");
                            writeDATA.print(setLECTURAS.getLectura_act() + ",");
                            writeDATA.print(setLECTURAS.getAnomalia_1() + ",");
                            writeDATA.print(setLECTURAS.getAnomalia_2() + ",");
                            writeDATA.print(setLECTURAS.getCodigo_operario() + ",");
                            writeDATA.print(setLECTURAS.getVigencia() + ",");
                            writeDATA.print(setLECTURAS.getFecha() + ",");
                            writeDATA.print(setLECTURAS.getOrden_lectura() + ",");
                            writeDATA.print(setLECTURAS.getLeido() + ",");
                            writeDATA.print(setLECTURAS.getCalle() + ",");
                            writeDATA.print(setLECTURAS.getEdificio() + ",");
                            writeDATA.print(setLECTURAS.getSuplemento_casa() + ",");
                            writeDATA.print(setLECTURAS.getInterloc_comercial() + ",");
                            writeDATA.print(setLECTURAS.getApellido() + ",");
                            writeDATA.print(setLECTURAS.getNombre() + ",");
                            writeDATA.print(setLECTURAS.getClase_instalacion());
                            writeDATA.println();
                        }
                        writeDATA.close();
    // 3. IMPORTAR LISTA DE DATOS A LA BASE DE DATOS

                        //CREAR ARCHIVO DE COMANDOS CON LAS RUTAS DE LA BASE DE DATOS Y EL ARCHIVO
                        File RutaDB = new File("dbs\\LECTURAS");
                        File RutaCARPETA = new File("lib\\sqlite-tools");
                        File RutaCOMANDOS = new File("lib\\sqlite-tools\\comandos.txt");
                        PrintWriter writeCOMANDOS = new PrintWriter(RutaCOMANDOS); //PARA ESCRIBIR EL COMANDO CON LA RUTA DE LOS DATOS
                        //COMANDO (script)
                        writeCOMANDOS.println(".mode csv");
                        writeCOMANDOS.println(".open '" + RutaDB.getAbsolutePath() + "'");
                        writeCOMANDOS.println(".import '" + RutaDATA.getAbsolutePath() + "' DATOS");
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
    //4. REINICIAR LISTAS PARA LA FILTRACION
                        panelCODPOR.removeAll();
                        panelANOM1.removeAll();
                        panelCODOPE.removeAll();
                        panelVIG.removeAll();
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

    //METODOS PARA FILTRAR DATOS
    //FILTRAR CODIGO PORCION
    public void filCODPOR() {
        //PARA CONVERTIR LA LISTA EN CHECKBOXES
        listCodPor = new ArrayList<CLASE_codpor>();
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        String query = "SELECT codigo_porcion FROM DATOS";
        try {
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                CLASE_codpor codpor = new CLASE_codpor();
                codpor.setCodpor(rs.getString("codigo_porcion"));
                listCodPor.add(codpor);
            }
            con.close();
        } catch (Exception ex) {
        }
        listCodPor = listCodPor.stream().distinct().collect(Collectors.toList());
        listCodPor =  listCodPor.stream().sorted(Comparator.comparing(CLASE_codpor::getCodpor)).collect(Collectors.toList());

        //DISEÑO JFRAME
        JPanel panelCHECKBOX = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
        panelCHECKBOX.setLayout(new BoxLayout(panelCHECKBOX, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y

        JCheckBox[] CHBX; //NUEVO ARRAY DE CHECKBOX
        CHBX = new JCheckBox[listCodPor.size()]; //INICIALIZAR CHECKBOX CON EL TAMAÑO DE LA LISTA DE LOS ELEMENTOS

        //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
        for (int j = 0; j < listCodPor.size(); j++) {
            CHBX[j] = new JCheckBox(listCodPor.get(j).getCodpor());
            panelCHECKBOX.add(CHBX[j]);
        }

        JScrollPane scroll = new JScrollPane(panelCHECKBOX); //NUEVO SCROLLPANE PARA EL panelSCROLL
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
        scroll.setPreferredSize(new Dimension (160, 100)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

        //CONTENIDO
        //PANEL PARA LOS CHECKBOX
        JPanel panelSCROLL = new JPanel(); //NUEVO PANEL PARA EL SCROLL DE LOS ELEMENTOS
        panelSCROLL.add(scroll); //AGREGANDO AL PANEL EL SCROLL QUE CONTIENE EL PANEL DE LOS CHECKBOX
        //PANEL PARA EL BOTON
        JPanel panelBOTON = new JPanel(); //NUEVO PANEL PARA EL BOTON
        JButton btnFiltrar = new JButton("CONFIRMAR"); //NUEVO BOTON
        btnFiltrar.setPreferredSize(new Dimension(160, 30)); //ASIGNAR EL TAMAÑO DEL BOTON
        panelBOTON.add(btnFiltrar, BorderLayout.PAGE_END); //ASIGNAR AL PANEL EL BOTON

        //AGREGANDO AL PANEL PRINCIPAL LOS PANELES CON EL CONTENIDO
        panelCODPOR.add(panelSCROLL); //AGREGAR EL PANEL DE LOS CHECKBOXS AL PANEL PRINCIPAL
        panelCODPOR.add(panelBOTON, BorderLayout.PAGE_END); //AGREGAR EL PANEL DEL BOTON AL PANEL PRINCIPAL

        frameCODPOR.setUndecorated(true);
        frameCODPOR.setContentPane(panelCODPOR);
        frameCODPOR.setSize(183, 150);
        frameCODPOR.setLocation(621, 647);
        frameCODPOR.setAlwaysOnTop(true);
        frameCODPOR.setVisible(true);

        //ACTION BUTTON
        btnFiltrar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String vacio = "()";
                String datos = "";
                queryFilCodPorcion = "(";
                int separadores = -1;

                for (int j = 0; j < listCodPor.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        separadores++;
                    }
                }

                for (int j = 0; j < listCodPor.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        datos = datos + CHBX[j].getText();
                        queryFilCodPorcion = queryFilCodPorcion + "codigo_porcion = '";
                        queryFilCodPorcion = queryFilCodPorcion + CHBX[j].getText();
                        queryFilCodPorcion = queryFilCodPorcion + "'";
                        if (0 < separadores) {
                            separadores--;
                            datos = datos + " ";
                            queryFilCodPorcion = queryFilCodPorcion + " OR ";
                        }
                    }
                }
                queryFilCodPorcion = queryFilCodPorcion + ")";
                if (queryFilCodPorcion.equals(vacio)) {
                    queryFilCodPorcion = "";
                    contCodPor = contCodPor + 1;
                    contValCodPor = 0;
                    frameCODPOR.dispose();
                } else {
                    frameCODPOR.dispose();
                    contCodPor = contCodPor + 1;
                    contValCodPor = 1;
                    JOptionPane.showMessageDialog(null, "SE FILTRARA LOS DATOS " + datos);
                }
            }
        });

    }
    //FILTRAR ANOMALIA 1
    public void filANOM1() {
        //PARA CONVERTIR LA LISTA EN CHECKBOXES
        listANOM1 = new ArrayList<CLASE_anom1>();
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        String query = "SELECT anomalia_1 FROM DATOS";
        try {
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                CLASE_anom1 anom1 = new CLASE_anom1();
                anom1.setAnom1(rs.getString("anomalia_1"));
                listANOM1.add(anom1);
            }
            con.close();
        } catch (Exception ex) {
        }
        listANOM1 = listANOM1.stream().distinct().collect(Collectors.toList());

        List<Integer> newListANOM1 = listANOM1.stream()
                .filter(anom1 -> !anom1.getAnom1().equals("")).map(anom1 -> Integer.parseInt(anom1.getAnom1()))
                .sorted(Comparator.comparing(getAnom1 -> getAnom1))
                .collect(Collectors.toList());

        //DISEÑO JFRAME
        JPanel panelCHECKBOX = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
        panelCHECKBOX.setLayout(new BoxLayout(panelCHECKBOX, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y

        JCheckBox[] CHBX; //NUEVO ARRAY DE CHECKBOX
        CHBX = new JCheckBox[newListANOM1.size()+1]; //INICIALIZAR CHECKBOX CON EL TAMAÑO DE LA LISTA DE LOS ELEMENTOS
        //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
        for (int j = 0; j < newListANOM1.size()+1; j++) {
            if (j == 0) {
                CHBX[j] = new JCheckBox("VACIOS");
                panelCHECKBOX.add(CHBX[j]);
            } else {
                CHBX[j] = new JCheckBox(newListANOM1.get(j-1).toString());
                panelCHECKBOX.add(CHBX[j]);
            }
        }

        JScrollPane scroll = new JScrollPane(panelCHECKBOX); //NUEVO SCROLLPANE PARA EL panelSCROLL
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
        scroll.setPreferredSize(new Dimension (124, 100)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

        //CONTENIDO
        //PANEL PARA LOS CHECKBOX
        JPanel panelSCROLL = new JPanel(); //NUEVO PANEL PARA EL SCROLL DE LOS ELEMENTOS
        panelSCROLL.add(scroll); //AGREGANDO AL PANEL EL SCROLL QUE CONTIENE EL PANEL DE LOS CHECKBOX
        //PANEL PARA EL BOTON
        JPanel panelBOTON = new JPanel(); //NUEVO PANEL PARA EL BOTON
        JButton btnFiltrar = new JButton("CONFIRMAR"); //NUEVO BOTON
        btnFiltrar.setPreferredSize(new Dimension(124, 30)); //ASIGNAR EL TAMAÑO DEL BOTON
        panelBOTON.add(btnFiltrar, BorderLayout.PAGE_END); //ASIGNAR AL PANEL EL BOTON

        //AGREGANDO AL PANEL PRINCIPAL LOS PANELES CON EL CONTENIDO
        panelANOM1.add(panelSCROLL); //AGREGAR EL PANEL SCROLL CON LOS CHECKBOXS AL PANEL PRINCIPAL
        panelANOM1.add(panelBOTON, BorderLayout.PAGE_END); //AGREGAR EL PANEL DEL BOTON AL PANEL PRINCIPAL

        frameANOM1.setUndecorated(true);
        frameANOM1.setContentPane(panelANOM1);
        frameANOM1.setSize(147, 150);
        frameANOM1.setLocation(816, 647);
        frameANOM1.setAlwaysOnTop(true);
        frameANOM1.setVisible(true);

        //ACTION BUTTON
        btnFiltrar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String vacio = "()";
                String datos = "";
                queryFilAnomalia1 = "(";
                int separadores = -1;

                for (int j = 0; j < newListANOM1.size()+1; j++) {
                    if (CHBX[j].isSelected()) {
                        separadores++;
                    }
                }

                for (int j = 0; j < newListANOM1.size()+1; j++) {
                    if (CHBX[j].isSelected()) {
                        datos = datos + CHBX[j].getText();
                        queryFilAnomalia1 = queryFilAnomalia1 + "anomalia_1 = '";
                        if (CHBX[j].getText() != "VACIOS") {
                            queryFilAnomalia1 = queryFilAnomalia1 + CHBX[j].getText();
                        }
                        queryFilAnomalia1 = queryFilAnomalia1 + "'";
                        if (0 < separadores) {
                            separadores--;
                            datos = datos + " ";
                            queryFilAnomalia1 = queryFilAnomalia1 + " OR ";
                        }
                    }
                }
                queryFilAnomalia1 = queryFilAnomalia1 + ")";
                if (queryFilAnomalia1.equals(vacio)) {
                    queryFilAnomalia1 = "";
                    contANOM1 = contANOM1 + 1;
                    contValANOM1 = 0;
                    frameANOM1.dispose();
                } else {
                    frameANOM1.dispose();
                    contANOM1 = contANOM1 + 1;
                    contValANOM1 = 1;
                    JOptionPane.showMessageDialog(null, "SE FILTRARA LOS DATOS " + datos);
                }
            }
        });
    }
    //FILTRAR CODIGO OPERARIO
    public void filCODOPE() {
        //PARA CONVERTIR LA LISTA EN CHECKBOXES
        listCodOpe = new ArrayList<CLASE_codope>();
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        String query = "SELECT codigo_operario FROM DATOS";
        try {
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                CLASE_codope codope = new CLASE_codope();
                codope.setCodope(rs.getString("codigo_operario"));
                listCodOpe.add(codope);
            }
            con.close();
        } catch (Exception ex) {
        }
        listCodOpe = listCodOpe.stream().distinct().collect(Collectors.toList());
        listCodOpe =  listCodOpe.stream().sorted(Comparator.comparing(CLASE_codope::getCodope)).collect(Collectors.toList());

        //DISEÑO JFRAME
        JPanel panelCHECKBOX = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
        panelCHECKBOX.setLayout(new BoxLayout(panelCHECKBOX, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y

        JCheckBox[] CHBX; //NUEVO ARRAY DE CHECKBOX
        CHBX = new JCheckBox[listCodOpe.size()]; //INICIALIZAR CHECKBOX CON EL TAMAÑO DE LA LISTA DE LOS ELEMENTOS

        //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
        for (int j = 0; j < listCodOpe.size(); j++) {
            CHBX[j] = new JCheckBox(listCodOpe.get(j).getCodope());
            panelCHECKBOX.add(CHBX[j]);
        }

        JScrollPane scroll = new JScrollPane(panelCHECKBOX); //NUEVO SCROLLPANE PARA EL panelSCROLL
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
        scroll.setPreferredSize(new Dimension (166, 100)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

        //CONTENIDO
        //PANEL PARA LOS CHECKBOX
        JPanel panelSCROLL = new JPanel(); //NUEVO PANEL PARA EL SCROLL DE LOS ELEMENTOS
        panelSCROLL.add(scroll); //AGREGANDO AL PANEL EL SCROLL QUE CONTIENE EL PANEL DE LOS CHECKBOX
        //PANEL PARA EL BOTON
        JPanel panelBOTON = new JPanel(); //NUEVO PANEL PARA EL BOTON
        JButton btnFiltrar = new JButton("CONFIRMAR"); //NUEVO BOTON
        btnFiltrar.setPreferredSize(new Dimension(166, 30)); //ASIGNAR EL TAMAÑO DEL BOTON
        panelBOTON.add(btnFiltrar, BorderLayout.PAGE_END); //ASIGNAR AL PANEL EL BOTON

        //AGREGANDO AL PANEL PRINCIPAL LOS PANELES CON EL CONTENIDO
        panelCODOPE.add(panelSCROLL); //AGREGAR EL PANEL DE LOS CHECKBOXS AL PANEL PRINCIPAL
        panelCODOPE.add(panelBOTON, BorderLayout.PAGE_END); //AGREGAR EL PANEL DEL BOTON AL PANEL PRINCIPAL

        frameCODOPE.setUndecorated(true);
        frameCODOPE.setContentPane(panelCODOPE);
        frameCODOPE.setSize(189, 150);
        frameCODOPE.setLocation(975, 647);
        frameCODOPE.setAlwaysOnTop(true);
        frameCODOPE.setVisible(true);

        //ACTION BUTTON
        btnFiltrar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String vacio = "()";
                String datos = "";
                queryFilCodOperario = "(";
                int separadores = -1;

                for (int j = 0; j < listCodOpe.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        separadores++;
                    }
                }

                for (int j = 0; j < listCodOpe.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        datos = datos + CHBX[j].getText();
                        queryFilCodOperario = queryFilCodOperario + "codigo_operario = '";
                        queryFilCodOperario = queryFilCodOperario + CHBX[j].getText();
                        queryFilCodOperario = queryFilCodOperario + "'";
                        if (0 < separadores) {
                            separadores--;
                            datos = datos + " ";
                            queryFilCodOperario = queryFilCodOperario + " OR ";
                        }
                    }
                }
                queryFilCodOperario = queryFilCodOperario + ")";
                if (queryFilCodOperario.equals(vacio)) {
                    queryFilCodOperario = "";
                    contCodOpe = contCodOpe + 1;
                    contValCodOpe = 0;
                    frameCODOPE.dispose();
                } else {
                    frameCODOPE.dispose();
                    contCodOpe = contCodOpe + 1;
                    contValCodOpe = 1;
                    JOptionPane.showMessageDialog(null, "SE FILTRARA LOS DATOS " + datos);
                }
            }
        });
    }
    //FILTRAR VIGENCIA
    public void filVIG() {
        //PARA CONVERTIR LA LISTA EN CHECKBOXES
        listVig = new ArrayList<CLASE_vig>();
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        String query = "SELECT vigencia FROM DATOS";
        try {
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                CLASE_vig vig = new CLASE_vig();
                vig.setVig(rs.getString("vigencia"));
                listVig.add(vig);
            }
            con.close();
        } catch (Exception ex) {
        }
        listVig = listVig.stream().distinct().collect(Collectors.toList());
        listVig =  listVig.stream().sorted(Comparator.comparing(CLASE_vig::getVig).reversed()).collect(Collectors.toList());

        //DISEÑO JFRAME
        JPanel panelCHECKBOX = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
        panelCHECKBOX.setLayout(new BoxLayout(panelCHECKBOX, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y

        JCheckBox[] CHBX; //NUEVO ARRAY DE CHECKBOX
        CHBX = new JCheckBox[listVig.size()]; //INICIALIZAR CHECKBOX CON EL TAMAÑO DE LA LISTA DE LOS ELEMENTOS

        //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
        for (int j = 0; j < listVig.size(); j++) {
            CHBX[j] = new JCheckBox(listVig.get(j).getVig());
            panelCHECKBOX.add(CHBX[j]);
        }

        JScrollPane scroll = new JScrollPane(panelCHECKBOX); //NUEVO SCROLLPANE PARA EL panelSCROLL
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
        scroll.setPreferredSize(new Dimension (106, 100)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

        //CONTENIDO
        //PANEL PARA LOS CHECKBOX
        JPanel panelSCROLL = new JPanel(); //NUEVO PANEL PARA EL SCROLL DE LOS ELEMENTOS
        panelSCROLL.add(scroll); //AGREGANDO AL PANEL EL SCROLL QUE CONTIENE EL PANEL DE LOS CHECKBOX
        //PANEL PARA EL BOTON
        JPanel panelBOTON = new JPanel(); //NUEVO PANEL PARA EL BOTON
        JButton btnFiltrar = new JButton("CONFIRMAR"); //NUEVO BOTON
        btnFiltrar.setPreferredSize(new Dimension(106, 30)); //ASIGNAR EL TAMAÑO DEL BOTON
        panelBOTON.add(btnFiltrar, BorderLayout.PAGE_END); //ASIGNAR AL PANEL EL BOTON

        //AGREGANDO AL PANEL PRINCIPAL LOS PANELES CON EL CONTENIDO
        panelVIG.add(panelSCROLL); //AGREGAR EL PANEL DE LOS CHECKBOXS AL PANEL PRINCIPAL
        panelVIG.add(panelBOTON, BorderLayout.PAGE_END); //AGREGAR EL PANEL DEL BOTON AL PANEL PRINCIPAL

        frameVIG.setUndecorated(true);
        frameVIG.setContentPane(panelVIG);
        frameVIG.setSize(129, 150);
        frameVIG.setLocation(1176, 647);
        frameVIG.setAlwaysOnTop(true);
        frameVIG.setVisible(true);

        //ACTION BUTTON
        btnFiltrar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String vacio = "()";
                String datos = "";
                queryFilVig = "(";
                int separadores = -1;

                for (int j = 0; j < listVig.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        separadores++;
                    }
                }

                for (int j = 0; j < listVig.size(); j++) {
                    if (CHBX[j].isSelected()) {
                        datos = datos + CHBX[j].getText();
                        queryFilVig = queryFilVig + "vigencia = '";
                        queryFilVig = queryFilVig + CHBX[j].getText();
                        queryFilVig = queryFilVig + "'";
                        if (0 < separadores) {
                            separadores--;
                            datos = datos + " ";
                            queryFilVig = queryFilVig + " OR ";
                        }
                    }
                }
                queryFilVig = queryFilVig + ")";
                if (queryFilVig.equals(vacio)) {
                    queryFilVig = "";
                    contVig = contVig + 1;
                    contValVig = 0;
                    frameVIG.dispose();
                } else {
                    frameVIG.dispose();
                    contVig = contVig + 1;
                    contValVig = 1;
                    JOptionPane.showMessageDialog(null, "SE FILTRARA LOS DATOS " + datos);
                }
            }
        });
    }
    //FILTRAR FECHA
    public void filFEC() {
        //DISEÑO JFRAME
        JLabel labelDESDE = new JLabel("DESDE");
        JDateChooser chooserDESDE = new JDateChooser();
        JLabel labelHASTA = new JLabel("HASTA");
        JDateChooser chooserHASTA = new JDateChooser();

        JPanel panelFECHA = new JPanel(new BorderLayout());
        panelFECHA.setLayout(new GridLayout(0, 1));

        JPanel textoDESDE = new JPanel(new BorderLayout());
        textoDESDE.setLayout(new GridBagLayout());
        textoDESDE.add(labelDESDE);

        JPanel choserDESDE = new JPanel(new BorderLayout());
        choserDESDE.setLayout(new GridLayout());
        choserDESDE.add(chooserDESDE);

        JPanel textoHASTA = new JPanel(new BorderLayout());
        textoHASTA.setLayout(new GridBagLayout());
        textoHASTA.add(labelHASTA);

        JPanel choserHASTA = new JPanel(new BorderLayout());
        choserHASTA.setLayout(new GridLayout());
        choserHASTA.add(chooserHASTA);

        //PANEL FECHA CON TODOS LOS COMPONENTES
        panelFECHA.add(textoDESDE);
        panelFECHA.add(choserDESDE);
        panelFECHA.add(textoHASTA);
        panelFECHA.add(choserHASTA);

        //PANEL PARA EL BOTON
        JPanel panelBOTON = new JPanel(); //NUEVO PANEL PARA EL BOTON
        JButton btnFiltrar = new JButton("CONFIRMAR"); //NUEVO BOTON
        btnFiltrar.setPreferredSize(new Dimension(109, 30)); //ASIGNAR EL TAMAÑO DEL BOTON
        panelBOTON.add(btnFiltrar, BorderLayout.PAGE_END); //ASIGNAR AL PANEL EL BOTON

        panelFEC.add(panelFECHA, BorderLayout.CENTER);
        panelFEC.add(panelBOTON, BorderLayout.PAGE_END);

        frameFEC.setUndecorated(true);
        frameFEC.setContentPane(panelFEC);
        frameFEC.setSize(109, 150);
        frameFEC.setAlwaysOnTop(true);
        frameFEC.setLocation(1317, 647);
        frameFEC.setVisible(true);

        //ACTION BUTTON
        btnFiltrar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (chooserDESDE.getDate() == null && chooserHASTA.getDate() == null) {
                    frameFEC.dispose();
                    queryFilFec = "";
                    contValFec = 0;
                    contFec = contFec + 1;
                } else if (chooserDESDE.getDate() != null && chooserHASTA.getDate() != null) {
                    Date dateDesde = chooserDESDE.getDate();
                    Date dateHasta = chooserHASTA.getDate();

                    String strDateDesde = DateFormat.getDateInstance().format(dateDesde);
                    String strDateHasta = DateFormat.getDateInstance().format(dateHasta);

                    if (dateDesde.compareTo(dateHasta) == -1 || dateDesde.compareTo(dateHasta) == 0) {
                        try {
                            Calendar gregorianCalendar = new GregorianCalendar();
                            DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                            Locale locale = new Locale("es", "EC");
                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd", locale);

                            Date date1 = dateFormat.parse(strDateDesde);
                            gregorianCalendar.setTime(date1);
                            rangoDesde = simpleDateFormat.format(date1);

                            Date date2 = dateFormat.parse(strDateHasta);
                            gregorianCalendar.setTime(date2);
                            rangoHasta = simpleDateFormat.format(date2);

                            queryFilFec = "(fecha BETWEEN '" + rangoDesde + "' AND '" + rangoHasta + ",')";

                            frameFEC.dispose();
                            contFec = contFec + 1;
                            contValFec = 1;
                            JOptionPane.showMessageDialog(null, "SE FILTRARA LOS DATOS DESDE " + strDateDesde + " HASTA " + strDateHasta);

                        } catch (ParseException ex) {
                            throw new RuntimeException(ex);
                        }
                    } else {
                        contValFec = 0;
                        JOptionPane.showMessageDialog(null, "ERROR: SELECCIONE UN RANGO DE FECHAS VALIDAS");
                    }

                } else {
                    contValFec = 0;
                    JOptionPane.showMessageDialog(null, "ERROR: NO PUEDEN HABER CAMPOS VACIOS");
                }
            }
        });
    }

    //3. METODOS PARA EXPORTAR
    public void EXPORT(){
        new Thread(()-> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
        new Thread(()-> ANOMALIAS()).start(); //INICIAR TAREA ANOMALIAS
    }

    public void ANOMALIAS() {
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION

        //PRIMER PASO: UNA LISTA DE TODAS LAS ANOMALIAS Y DE TODAS LAS VIGENCIAS EXISTENTES
        List Anomalias = new ArrayList<Integer>();
        List Descripcion = new ArrayList<String>();
        Vigencias = new ArrayList<VIGENCIAS>();
        AnomaliasXVigencia = new ArrayList<ANOMXVIG>();
        List AXV = new ArrayList<String>();
        String AxV = "";

        try {
            //ANOMALIAS
            for (int i = 4; i <= 30; i++) {
                if (i == 22) {
                    i += 1;
                }
                Anomalias.add(i);
            }
            //DESCRIPCION
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

            //VIGENCIAS
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM DATOS ORDER BY vigencia");
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
                    PreparedStatement psANOMXVIG = con.prepareStatement("SELECT count(anomalia_1) as \"ANOMXVIG\" FROM DATOS WHERE ((anomalia_1 != \"\") AND (anomalia_1 =" + i + ") AND vigencia = '" + Vigencias.get(j).getVigencia() + "')");
                    ResultSet rsANOMXVIG = psANOMXVIG.executeQuery();
                    ANOMXVIG AnomXVig = new ANOMXVIG();
                    AnomXVig.setAnomXVig(rsANOMXVIG.getString("ANOMXVIG"));
                    AnomaliasXVigencia.add(AnomXVig);
                }

                int contador = 0;
                for (ANOMXVIG model : AnomaliasXVigencia) {
                    contador += 1;
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
            con.close();

            File csvFile = new File("files\\ANOMALIAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter write = new PrintWriter(csvFile); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "ANOMALIA,DESCRIPCION,";

            int separadores = -1;

            for (int j = 0; j < Vigencias.size(); j++) {
                separadores++;
            }

            for (VIGENCIAS Vigencia : Vigencias) {
                estructura += (Vigencia.getVigencia());
                if (0 < separadores) {
                    separadores--;
                    estructura += ",";
                }
            }
            write.println(estructura);

            for (int j = 0; j < Anomalias.size(); j++) {
                write.print(Anomalias.get(j) + ",");
                write.print(Descripcion.get(j));
                if (j  < Anomalias.size()){
                    write.print("," + AXV.get(j));
                } else {
                    write.print(",");
                }
                write.println();
            }
            write.close();

            //CONVERTIR EN EXCEL
            Workbook wbANOMALIAS = new Workbook("files\\ANOMALIAS.csv"); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS

            Worksheet wsANOMALIAS = wbANOMALIAS.getWorksheets().get(0); //NUEVA HOJA DE ANOMALIAS PARA EL LIBRO DE ANOMALIAS
            //ASIGNAR CELDAS CON UN TAMAÑO DEFINIDO
            Cells cells = wsANOMALIAS.getCells();
            cells.setColumnWidth(0, 10);
            cells.setColumnWidth(1, 30);
            //ALINEAR CELDAS A LA IZQUIERDA
            Style st1 = wbANOMALIAS.createStyle();
            st1.setHorizontalAlignment(TextAlignmentType.LEFT);
            st1.setVerticalAlignment(TextAlignmentType.CENTER);
            StyleFlag flag = new StyleFlag();
            flag.setAlignments(true);
            Range rng = wsANOMALIAS.getCells().createRange("A1:B28");
            rng.applyStyle(st1, flag);
            //ALINEAR CELDAS A LA DERECHA
            Style st2 = wbANOMALIAS.createStyle();
            st2.setHorizontalAlignment(TextAlignmentType.RIGHT);
            st2.setVerticalAlignment(TextAlignmentType.CENTER);
            StyleFlag flag2 = new StyleFlag();
            flag2.setAlignments(true);
            Range rng2 = wsANOMALIAS.getCells().createRange("C1:AZ28");
            rng2.applyStyle(st2, flag2);
            //COLOREAR CELDA B25 Y B26 DESCRIPCIONES DE LAS ANOMALIAS
            Cell B25 = wsANOMALIAS.getCells().get("B25");
            Style stB25 = B25.getStyle();
            stB25.setPattern(BackgroundType.SOLID);
            stB25.setForegroundColor(com.aspose.cells.Color.getYellow());
            B25.setStyle(stB25);

            Cell B26 = wsANOMALIAS.getCells().get("B26");
            Style stB26 = B26.getStyle();
            stB26.setPattern(BackgroundType.SOLID);
            stB26.setForegroundColor(com.aspose.cells.Color.getYellow());
            B26.setStyle(stB26);

            wbANOMALIAS.save("files\\INFORME.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            csvFile.delete(); //ELIMINAR ARCHIVO DE Datos.csv

            fCargar.dispose(); // CERRAR PANTALLA DE CARGA
            JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE EL INFORME");
            Runtime.getRuntime().exec("cmd /c start cmd.exe /K \" start " + ARCHIVOS.getAbsolutePath() + " && exit");

        } catch (Exception ex) {
        }
    }

    //3.2 METODO QUE EXPORTA DATIS FILTRADOS A xlsx
    public void EXPORTARfiltrados () {
        DATOSdb = new ArrayList<getLECTURAS>();
        conexion_lectura sql = new conexion_lectura(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION

        String queryFil = "";
        int contFiltraciones = 0;
        contFiltraciones = (contValCodPor + contValANOM1 + contValCodOpe + contValVig + contValFec);

        if (contFiltraciones != 0) {
            if (contValCodPor == 1) {
                queryFil = queryFil + queryFilCodPorcion;
                if (1 < contFiltraciones) {
                    contFiltraciones--;
                    queryFil = queryFil + " AND ";
                }
            }
            if (contValANOM1 == 1) {
                queryFil = queryFil + queryFilAnomalia1;
                if (1 < contFiltraciones) {
                    contFiltraciones--;
                    queryFil = queryFil + " AND ";
                }
            }
            if (contValCodOpe == 1) {
                queryFil = queryFil + queryFilCodOperario;
                if (1 < contFiltraciones) {
                    contFiltraciones--;
                    queryFil = queryFil + " AND ";
                }
            }
            if (contValVig == 1) {
                queryFil = queryFil + queryFilVig;
                if (1 < contFiltraciones) {
                    contFiltraciones--;
                    queryFil = queryFil + " AND ";
                }
            }
            if (contValFec == 1) {
                queryFil = queryFil + queryFilFec;
                if (1 < contFiltraciones) {
                    contFiltraciones--;
                    queryFil = queryFil + " AND ";
                }
            }
            try {
                new Thread(()-> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
                PreparedStatement ps = con.prepareStatement("SELECT * FROM DATOS WHERE " + queryFil);
                ResultSet rs = ps.executeQuery();
                while (rs.next()) {
                    getLECTURAS dbDATOS = new getLECTURAS();
                    dbDATOS.setCodigo_porcion(rs.getString("codigo_porcion"));
                    dbDATOS.setUni_lectura(rs.getString("uni_lectura"));
                    dbDATOS.setDoc_lectura(rs.getString("doc_lectura"));
                    dbDATOS.setCuenta_contrato(rs.getString("cuenta_contrato"));
                    dbDATOS.setMedidor(rs.getString("medidor"));
                    dbDATOS.setLectura_ant(rs.getString("lectura_ant"));
                    dbDATOS.setLectura_act(rs.getString("lectura_act"));
                    dbDATOS.setAnomalia_1(rs.getString("anomalia_1"));
                    dbDATOS.setAnomalia_2(rs.getString("anomalia_2"));
                    dbDATOS.setCodigo_operario(rs.getString("codigo_operario"));
                    dbDATOS.setVigencia(rs.getString("vigencia"));
                    dbDATOS.setFecha(rs.getString("fecha"));
                    dbDATOS.setOrden_lectura(rs.getString("orden_lectura"));
                    dbDATOS.setLeido(rs.getString("leido"));
                    dbDATOS.setCalle(rs.getString("calle"));
                    dbDATOS.setEdificio(rs.getString("edificio"));
                    dbDATOS.setSuplemento_casa(rs.getString("suplemento_casa"));
                    dbDATOS.setInterloc_comercial(rs.getString("interloc_comercial"));
                    dbDATOS.setApellido(rs.getString("apellido"));
                    dbDATOS.setNombre(rs.getString("nombre"));
                    dbDATOS.setClase_instalacion(rs.getString("clase_instalacion"));
                    DATOSdb.add(dbDATOS);
                }
                con.close();

                File csvFile = new File("files\\Datos.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
                PrintWriter write = new PrintWriter(csvFile); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO
                String estructura = "CODIGO_PORCION,UNI_LECTURA,DOC_LECTURA,CUENTA_CONTRATO,MEDIDOR,LEC_ANTERIOR,LEC_ACTUAL,ANOMALIA_1,ANOMALIA_2,CODIGO_OPERARIO,VIGENCIA,FECHA,ORDEN_LECTURA,LEIDO,CALLE,EDIFICIO,SUPLEMENTO_CASA,INTERLOC_COM,APELLIDO,NOMBRE,CLASE_INSTALA";
                write.println(estructura);

                for (getLECTURAS datos : DATOSdb) {
                    write.print(datos.getCodigo_porcion() + ",");
                    write.print(datos.getUni_lectura() + ",");
                    write.print(datos.getDoc_lectura() + ",");
                    write.print(datos.getCuenta_contrato() + ",");
                    write.print(datos.getMedidor() + ",");
                    write.print(datos.getLectura_ant() + ",");
                    write.print(datos.getLectura_act() + ",");
                    write.print(datos.getAnomalia_1() + ",");
                    write.print(datos.getAnomalia_2() + ",");
                    write.print(datos.getCodigo_operario() + ",");
                    write.print(datos.getVigencia() + ",");
                    write.print(datos.getFecha() + ",");
                    write.print(datos.getOrden_lectura() + ",");
                    write.print(datos.getLeido() + ",");
                    write.print(datos.getCalle() + ",");
                    write.print(datos.getEdificio() + ",");
                    write.print(datos.getSuplemento_casa() + ",");
                    write.print(datos.getInterloc_comercial() + ",");
                    write.print(datos.getApellido() + ",");
                    write.print(datos.getNombre() + ",");
                    write.print(datos.getClase_instalacion());
                    write.println();
                }
                write.close();

                Workbook wbCSV = new Workbook("files\\Datos.csv"); //NUEVO LIBRO DEL ARCHIVO Datos
                wbCSV.save("files\\Filtrados.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
                csvFile.delete(); //ELIMINAR ARCHIVO DE Datos.csv

                fCargar.dispose();
                JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE " + DATOSdb.size() + " REGISTROS");
                Runtime.getRuntime().exec("cmd /c start cmd.exe /K \" start " + ARCHIVOS.getAbsolutePath() + " && exit");

                } catch (Exception ex) {
                }

            } else {
                JOptionPane.showMessageDialog(null,"ERROR: NO SE HA SELECCIONADO NINGUN FILTRO");
            }

    }

    //MAIN QUE ARRANCA EL PROGRAMA
    public static void main(String[] args) {
        new PROGRAMA();
    }
}


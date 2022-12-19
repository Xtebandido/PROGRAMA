package Principal;
//CLASES Y LIBRERIAS IMPORTADAS
import Conexion.DATABASE;
import Modelo.*;
import com.aspose.cells.*;
import com.csvreader.CsvReader;
import jnafilechooser.api.JnaFileChooser;
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

public class PROGRAMA extends JFrame {
    //VARIABLES DE PROGRAMA
    JPanel mainPanel; //PANEL PRINCIPAL
    //------LOADING------
    JDialog dialog; //DIALOGO QUE CONTIENE LA PANTALLA DE CARGA

    //[LECTURAS]
    //------INSERTAR-----
    JPanel jpIMPORT; //PANEL IMPORTAR
    JButton btnSELECT; //SELECCIONAR    ->  BOTON SELECCIONAR ARCHIVO
    File file = null; //SELECCIONAR    ->  ARCHIVO DONDE SE GUARDARA EL ARCHIVO SELECCIONADO
    JTextField jtxtPATH; //SELECCIONAR    ->  JTEXTFIELD CON EL DATO DE LA RUTA DEL ARCHIVO XLSX SELECCIONADO
    String PATH = ""; //IMPORTAR    ->  STRING QUE TIENE EL DATO DE LA RUTA DEL ARCHIVO SELECCIONADO PARA IMPORTAR
    JButton btnIMPORT; //IMPORTAR    ->  BOTON IMPORTAR
    //--------FILTRAR--------
    JPanel jpFILTER_LEC; //PANEL FILTRAR
    List<String> Porciones; //LISTA CON LOS DATOS TIPO STRING
    JButton btnFILTER_CODPOR; //BOTON PARA FILTRAR LA PORCION
    JPanel jpSCROLL_CODPOR = new JPanel(); //PANEL DONDE SE ENCUENTRA LOS CHECKBOX CON DESPLAZAMIENTOS
    JPopupMenu puMENU_CODPOR = new JPopupMenu(); //POPUPMENU CON EL panelSCROLL
    JCheckBox[] CHBX_CODPOR; //CHECKBOXS CON LOS DATOS
    List<String> Operarios; //LISTA CON LOS DATOS TIPO STRING
    JButton btnFILTER_CODOPE; //BOTON PARA FILTRAR LA PORCION
    JPanel jpSCROLL_CODOPE = new JPanel(); //PANEL DONDE SE ENCUENTRA LOS CHECKBOX CON DESPLAZAMIENTOS
    JPopupMenu puMENU_CODOPE = new JPopupMenu(); //POPUPMENU CON EL panelSCROLL
    JCheckBox[] CHBX_CODOPE; //CHECKBOXS CON LOS DATOS
    List<String> Vigencias; //LISTA CON LOS DATOS TIPO STRING
    JButton btnFILTER_VIG; //BOTON PARA FILTRAR LA VIGENCIA
    JPanel jpSCROLL_VIG = new JPanel(); //PANEL DONDE SE ENCUENTRA LOS CHECKBOX CON DESPLAZAMIENTOS
    JPopupMenu puMENU_VIG = new JPopupMenu(); //POPUPMENU CON EL panelSCROLL
    JCheckBox[] CHBX_VIG; //CHECKBOXS CON LOS DATOS
    //--------EXPORTAR--------
    JPanel jpEXPORT; //PANEL DE EXPORTAR DENTRO DEL PANEL DE LECTURAS
    JButton btnEXPORT; // BOTON PARA EXPORTAR TODOS LOS DATOS

    //----VALIDAR---
    int INITprogram;
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

        INIT();

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
                    JOptionPane.showMessageDialog(null, "SELECCIONE UN ARCHIVO", "", JOptionPane.INFORMATION_MESSAGE);
                }
            }
        });

        //BOTON FILTRAR PORCION
        btnFILTER_CODPOR.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                puMENU_CODPOR.show(btnFILTER_CODPOR, 0, btnFILTER_CODPOR.getHeight());
            }
        });
        //BOTON FILTRAR OPERARIO
        btnFILTER_CODOPE.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                puMENU_CODOPE.show(btnFILTER_CODOPE, 0, btnFILTER_CODOPE.getHeight());
            }
        });
        //BOTON FILTRAR VIGENCIA
        btnFILTER_VIG.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                puMENU_VIG.show(btnFILTER_VIG, 0, btnFILTER_VIG.getHeight());
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
        if (INITprogram != 0) {
            panelLOAD.add(new JLabel("CARGANDO REGISTROS... ESTO PODRIA TOMAR UNOS MINUTOS"), BorderLayout.PAGE_START); //AÑADIR UN LABEL AL INICIO DEL PANEL
        }
        if (INITprogram == 0) {
            panelLOAD.add(new JLabel("CARGANDO PROGRAMA..."), BorderLayout.PAGE_START); //AÑADIR UN LABEL AL INICIO DEL PANEL
        }

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

    public void INIT(){
        if (INITprogram == 0) {
            new Thread(() -> LOADING()).start();
        }
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //VIGENCIAS
            List<String> Vigencias = new ArrayList<String>(); //NUEVA LISTA DE VIGENCIAS
            PreparedStatement psVigencia = con.prepareStatement("SELECT DISTINCT vigencia FROM LECTURAS ORDER BY vigencia"); //QUERY
            ResultSet rsVigencia = psVigencia.executeQuery(); //RESULTADOS DE LA CONSULTA
            while (rsVigencia.next()) {
                String VIG = rsVigencia.getString("vigencia");
                Vigencias.add(VIG);
            }
            int vigINICIAL;
            int vigFINAL = Integer.parseInt(Vigencias.get(Vigencias.size()-1));

            for (int j = 0; j < Vigencias.size(); j++) {
                vigINICIAL = Integer.parseInt(Vigencias.get(j));
                if ((vigFINAL - vigINICIAL) >= 400) {
                    Statement delete = con.createStatement();
                    delete.executeUpdate("DELETE FROM LECTURAS WHERE vigencia = '" + vigINICIAL + "'");
                    Vigencias.remove(j);
                    j--;
                }
            }

            Collections.sort(Vigencias, Collections.reverseOrder()); //ORDENAR VIGENCIAS DE MENOR A MAYOR

            JPanel jpCHECK_VIG = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
            jpCHECK_VIG.setLayout(new BoxLayout(jpCHECK_VIG, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y
            //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
            CHBX_VIG = new JCheckBox[Vigencias.size()];; //NUEVO ARRAY DE CHECKBOX
            for (int j = 0; j < Vigencias.size(); j++) {
                CHBX_VIG[j] = new JCheckBox(Vigencias.get(j));
                jpCHECK_VIG.add(CHBX_VIG[j]);
            }

            JScrollPane jspVIG = new JScrollPane(jpCHECK_VIG); //NUEVO SCROLLPANE PARA EL panelSCROLL
            jspVIG.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
            jspVIG.getVerticalScrollBar().setUnitIncrement(145);
            jspVIG.setPreferredSize(new Dimension (258, 150)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

            jpSCROLL_VIG.add(jspVIG);
            puMENU_VIG.add(jpSCROLL_VIG);

            //OPERARIO
            List<String> Operarios = new ArrayList<String>(); //NUEVA LISTA DE VIGENCIAS
            PreparedStatement psOperarios = con.prepareStatement("SELECT DISTINCT codigo_operario FROM LECTURAS ORDER BY codigo_operario"); //QUERY
            ResultSet rsOperario = psOperarios.executeQuery(); //RESULTADOS DE LA CONSULTA
            while (rsOperario.next()) {
                String CODOPE = rsOperario.getString("codigo_operario");
                Operarios.add(CODOPE);
            }

            JPanel jpCHECK_CODOPE = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
            jpCHECK_CODOPE.setLayout(new BoxLayout(jpCHECK_CODOPE, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y
            //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
            CHBX_CODOPE = new JCheckBox[Operarios.size()];; //NUEVO ARRAY DE CHECKBOX
            for (int j = 0; j < Operarios.size(); j++) {
                CHBX_CODOPE[j] = new JCheckBox(Operarios.get(j));
                jpCHECK_CODOPE.add(CHBX_CODOPE[j]);
            }

            JScrollPane jspCODOPE = new JScrollPane(jpCHECK_CODOPE); //NUEVO SCROLLPANE PARA EL panelSCROLL
            jspCODOPE.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
            jspCODOPE.getVerticalScrollBar().setUnitIncrement(20);
            jspCODOPE.setPreferredSize(new Dimension (249, 150)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

            jpSCROLL_CODOPE.add(jspCODOPE);
            puMENU_CODOPE.add(jpSCROLL_CODOPE);

            //PORCION
            List<String> Porciones = new ArrayList<String>(); //NUEVA LISTA DE VIGENCIAS
            PreparedStatement psPorcion = con.prepareStatement("SELECT DISTINCT codigo_porcion FROM LECTURAS ORDER BY codigo_porcion"); //QUERY
            ResultSet rsPorcion = psPorcion.executeQuery(); //RESULTADOS DE LA CONSULTA
            while (rsPorcion.next()) {
                String CODPOR = rsPorcion.getString("codigo_porcion");
                Porciones.add(CODPOR);
            }

            JPanel jpCHECK_CODPOR = new JPanel(); //NUEVO PANEL PARA GUARDAR EL PANEL SCROLL
            jpCHECK_CODPOR.setLayout(new BoxLayout(jpCHECK_CODPOR, BoxLayout.Y_AXIS)); //ASIGNAR AL PANEL LOS ELEMENTOS DE FORMA VERTICAL EN EL EJE Y
            //CICLO QUE TOMA LOS ELEMENTOS DE LA LISTA Y LOS AGREGA AL CHECKBOX Y LOS ELEMENTOS SON AGREGADOS AL PANEL
            CHBX_CODPOR = new JCheckBox[Porciones.size()];; //NUEVO ARRAY DE CHECKBOX
            for (int j = 0; j < Porciones.size(); j++) {
                CHBX_CODPOR[j] = new JCheckBox(Porciones.get(j));
                jpCHECK_CODPOR.add(CHBX_CODPOR[j]);
            }

            JScrollPane jspCODPOR = new JScrollPane(jpCHECK_CODPOR); //NUEVO SCROLLPANE PARA EL panelSCROLL
            jspCODPOR.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); //ASIGNAR EL SCROLL VERTICAL
            jspCODPOR.getVerticalScrollBar().setUnitIncrement(20);
            jspCODPOR.setPreferredSize(new Dimension (244, 150)); //ASIGNAR EL TAMAÑO DE LA VENTANA DEL SCROLL

            jpSCROLL_CODPOR.add(jspCODPOR);
            puMENU_CODPOR.add(jpSCROLL_CODPOR);

            con.close();
        } catch (Exception ex) {
        }
        if (INITprogram == 0) {
            INITprogram++;
            dialog.dispose();
            setVisible(true);
        }
        if (INITprogram != 0) {
            dialog.dispose();
        }
    }

    //METODO SELECCIONAR ARCHIVO
    public void SELECTFILE() {
        JnaFileChooser fileChooser = new JnaFileChooser(); //FILECHOOSER PARA SELECCIONAR ARCHIVO
        fileChooser.addFilter("EXCEL", "xlsx", "xls"); //FILTRO PARA SELECCIONAR UNICAMENTE ARCHIVOS EXCEL
        boolean action = fileChooser.showOpenDialog(this);
        if (action){
            jtxtPATH.setText(fileChooser.getSelectedFile().toString());
            file = fileChooser.getCurrentDirectory();
            PATH = "" + fileChooser.getSelectedFile().toString();
        }
    }

    //METODO IMPORTAR
    public void FUN_IMPORT() {
        new Thread(() -> LOADING()).start(); //INICIAR TAREA DE PANTALLA DE CARGA
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

                        //RESETEAR LOS DATOS PARA FILTRAR Y GENERAR INFORME E INICIAR METODO INIT
                        jpSCROLL_CODPOR.removeAll();
                        puMENU_CODPOR.removeAll();
                        jpSCROLL_CODOPE.removeAll();
                        puMENU_CODOPE.removeAll();
                        jpSCROLL_VIG.removeAll();
                        puMENU_VIG.removeAll();
                        new Thread (()-> INIT()).run();

                        JOptionPane.showMessageDialog(null, "SE IMPORTO CORRECTAMENTE " + DATA.size() + " REGISTROS DE " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                        if (repetidosFinal.size() != 0) {
                            JOptionPane.showMessageDialog(null, "SE ENCONTRARON " + repetidosFinal.size() + " REGISTROS REPETIDOS EN " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                            File rutaARCHIVOS = new File("files");
                            Runtime.getRuntime().exec("cmd /c start " + rutaARCHIVOS.getAbsolutePath() + "\\REPETIDOS_" + fileNAME.getName() + " && exit");
                        }
                        //jtxtPATH.setText(null);
                        //PATH = "";
                    }
                } else {
                    dialog.dispose(); //CERRAR LOADING
                    JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LOS DATOS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
                }
            } else {
                dialog.dispose(); //CERRAR LOADING
                JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LA ESTRUCTURA DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR LA ESTRUCTURA DEL ARCHIVO
            }
        } catch (Exception e) {
            dialog.dispose(); //CERRAR LOADING
            File file = new File("files\\Importe.csv");
            file.delete();
            JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LAS FECHAS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
        }
    }

    //METODO VALIDAR SI EL INFORME SE ENCUENTRA ABIERTO, VALIDAR LOS AÑOS DE LA BASE DE DATOS QUE SEAN UNICAMENTE LOS ULTIMOS 4 AÑOS E INICIAR LAS TAREAS PARA REALIZAR EL INFORME
    public void CHECKING(){
        boolean fileOPEN = false;
        try {
            File file = new File("files\\INFORME.xlsx");
            //SI NO EXISTE CREAR UN ARCHIVO DE INFORME
            if (!file.exists()) {
                Workbook newWB = new Workbook(); //NUEVO LIBRO
                newWB.save("files\\INFORME.xlsx");
            }
            Workbook wb = new Workbook("files\\INFORME.xlsx"); //NUEVO LIBRO
            wb.save("files\\INFORME.xlsx"); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            file.delete();
        } catch (Exception e) {
            fileOPEN = true;
            dialog.dispose(); //CERRAR LOADING
            JOptionPane.showMessageDialog(null, "ERROR: EL INFORME NO PUEDE SER EXPORTADO PORQUE EL ARCHIVO SE ENCUENTRA ABIERTO. CIERRELO Y VUELVA A INTENTARLO", "", JOptionPane.INFORMATION_MESSAGE);
        }
        //SI EL ARCHIVO NO SE ENCUENTRA ABIERTO PROCEDER CON INICIAR LOS METODOS
        if (fileOPEN != true) {
            //RESETEAR LISTAS
            Porciones = new ArrayList<String>();
            Operarios = new ArrayList<String>();
            Vigencias = new ArrayList<String>();

            for (int j = 0; j < CHBX_CODPOR.length; j++) {
                if (CHBX_CODPOR[j].isSelected()){
                    Porciones.add(CHBX_CODPOR[j].getText());
                }
            }

            for (int j = 0; j < CHBX_CODOPE.length; j++) {
                if (CHBX_CODOPE[j].isSelected()){
                    Operarios.add(CHBX_CODOPE[j].getText());
                }
            }

            for (int j = 0; j < CHBX_VIG.length; j++) {
                if (CHBX_VIG[j].isSelected()){
                    Vigencias.add(CHBX_VIG[j].getText());
                }
            }

            if (Vigencias.size() == 0) {
                for (int j = 0; j < CHBX_VIG.length; j++) {
                    Vigencias.add(CHBX_VIG[j].getText());
                }
            }

            if (Porciones.size() == CHBX_CODPOR.length) {
                Porciones.clear();
            }

            if (Operarios.size() == CHBX_CODOPE.length) {
                Operarios.clear();
            }

            Collections.sort(Vigencias); //ORDENAR VIGENCIAS DE MENOR A MAYOR

            //INICIAR METODOS
            new Thread(() -> infoCONSUMO_0()).start();
            new Thread(() -> infoCONSUMOS_NEGATIVOS()).start();
            new Thread(() -> infoANOMALIAS()).start();
            new Thread(() -> infoANOMALIASxPORCION()).start();
            new Thread(() -> infoANOMALIASxOPERARIO()).start();
            //new Thread(() -> infoLECTURAS()).start();
        }
    }

    //METODO informe -> CONSUMO_0
    public void infoCONSUMO_0() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y PROCEDER AL ULTIMO METODO
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = " AND (";
            String nameOperarios = "";

            //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
            for (int j = 0; j < Operarios.size(); j++) {
                CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                nameOperarios += Operarios.get(j);
                if (j < (Operarios.size()-1)) {
                    CODOPE += " OR ";
                    nameOperarios += "-";
                }
            }
            nameOperarios = "\nOPERARIOS (" + nameOperarios + ")";
            CODOPE += ")";
            //SI NO HAY OPERARIOS FILTRADOS, VACIAR LOS STRINGS Y LISTAR TODAS LAS PORCIONES EN LA LISTA LOCAL
            if (Operarios.size() == 0) {
                CODOPE = "";
                nameOperarios = "";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA HACER ESTO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i)); //AGREGAR PORCIONES FILTRADAS A LA LISTA LOCAL
                query += "SELECT codigo_porcion"; //QUERY CON LAS PORCIONES FILTRADAS
                for (int j = 0; j < Vigencias.size(); j++) {
                    query += ", COUNT(*) FILTER(WHERE (lectura_act != '' AND lectura_ant != '') AND (lectura_act - lectura_ant = 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                }
                query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                if (i < (Porciones.size()-1)) {
                    query += " UNION ";
                }
            }
            //SI NO HAY NINGUNA PORCION FILTRADA HACER ESTO
            if (Porciones.size() == 0) {
                //CICLO QUE AGREGA TODAS LAS PORCIONES EN UNA LISTA LOCAL
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }
                query += "SELECT codigo_porcion"; //QUERY CON TODAS LAS PORCIONES
                for (int i = 0; i < Vigencias.size(); i++) {
                    query += ", COUNT(*) FILTER(WHERE (lectura_act != '' AND lectura_ant != '') AND (lectura_act - lectura_ant = 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + "'";
                }
                query += " FROM LECTURAS GROUP BY codigo_porcion";
            }

            //LISTAR VALORES EN UNA LISTA CON LISTAS DE VIGENCIAS
            List<List<String>> resultXvig = new ArrayList<List<String>>();
            for(int i = 0; i < Vigencias.size(); i++){
                resultXvig.add(new ArrayList<String>());
            }

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                //VIGENCIAS
                for (int j = 0; j < Vigencias.size(); j++) {
                    String anomXvig = rs.getString(Vigencias.get(j));
                    resultXvig.get(j).add(anomXvig);
                }
            }
            con.close(); //CERRAR CONEXION

            File fileCONSUMO_0 = new File("files\\CONSUMO_0.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeCONSUMO_0 = new PrintWriter(fileCONSUMO_0); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "PORCION,"; //ESTRUCTURA PRIMERA FILA
            //ESCRIBIR VIGENCIAS EN LA ESTRUCTURA - PRIMERA FILA
            for (int j = 0; j < Vigencias.size(); j++) {
                estructura += ("VIG" + Vigencias.get(j));
                if (j < (Vigencias.size()-1)) {
                    estructura += ",";
                }
            }
            writeCONSUMO_0.println(estructura);
            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                writeCONSUMO_0.print(porcionesLocal.get(i));
                for (int j = 0; j < Vigencias.size(); j++) {
                    writeCONSUMO_0.print("," + resultXvig.get(j).get(i));
                }
                writeCONSUMO_0.println();
            }
            writeCONSUMO_0.println("TOTAL");
            writeCONSUMO_0.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wbCONSUMO_0 = new Workbook("files\\CONSUMO_0.csv"); //NUEVO LIBRO
            Worksheet wsCONSUMO_0 = wbCONSUMO_0.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (wsCONSUMO_0.getCells().getCell(0,wsCONSUMO_0.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = wsCONSUMO_0.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA ANOM

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wbCONSUMO_0.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = wsCONSUMO_0.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsCONSUMO_0.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = wsCONSUMO_0.getCells().createRange("A1:A"+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR A LA DERECHA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsCONSUMO_0.getCells().createRange("B1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);

            Cell cell;
            for (int j = 0; j < Vigencias.size(); j++) {
                if (j < Vigencias.size()) {
                    String cellChar = (wsCONSUMO_0.getCells().getCell(0,(j+1)).getName()).replaceAll("1","");
                    cell = wsCONSUMO_0.getCells().get(cellChar + (porcionesLocal.size()+2));
                    cell.setFormula("=SUM(" + cellChar + "2:" + cellChar + (porcionesLocal.size()+1) + ")");
                }
            }
            //GRAFICAR
            //CREAR GRAFICA 'TOTAL ANOMALIAS SIN 18 Y 28' Y POSICIONARLA
            int idx1 = wsCONSUMO_0.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), 0, ((porcionesLocal.size()+3)+22), (Vigencias.size()+1));
            Chart ch1 = wsCONSUMO_0.getCharts().get(idx1);
            ch1.getTitle().setText("TOTAL CONSUMO 0 LECTURA " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
            ch1.getNSeries().add("A"+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
            ch1.getNSeries().setCategoryData("=B1:" + lastCell + "1"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
            ch1.getNSeries().get(0).setName("=A"+(porcionesLocal.size()+2)); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
            ch1.getNSeries().get(0).setValues("=B"+(porcionesLocal.size()+2)+":" + lastCell + +(porcionesLocal.size()+2)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
            ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
            ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

            wbCONSUMO_0.save("files\\CONSUMO_0.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileCONSUMO_0.delete(); //ELIMINAR ARCHIVO DE ANOMALIAS.csv

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> CONSUMOS_NEGATIVOS
    public void infoCONSUMOS_NEGATIVOS() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y PROCEDER AL ULTIMO METODO
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = " AND (";
            String nameOperarios = "";

            //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
            for (int j = 0; j < Operarios.size(); j++) {
                CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                nameOperarios += Operarios.get(j);
                if (j < (Operarios.size()-1)) {
                    CODOPE += " OR ";
                    nameOperarios += "-";
                }
            }
            nameOperarios = "\nOPERARIOS (" + nameOperarios + ")";
            CODOPE += ")";
            //SI NO HAY OPERARIOS FILTRADOS, VACIAR LOS STRINGS Y LISTAR TODAS LAS PORCIONES EN LA LISTA LOCAL
            if (Operarios.size() == 0) {
                CODOPE = "";
                nameOperarios = "";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA HACER ESTO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i)); //AGREGAR PORCIONES FILTRADAS A LA LISTA LOCAL
                query += "SELECT codigo_porcion"; //QUERY CON LAS PORCIONES FILTRADAS
                for (int j = 0; j < Vigencias.size(); j++) {
                    query += ", COUNT(*) FILTER(WHERE (lectura_act != '' AND lectura_ant != '') AND (lectura_act - lectura_ant < 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                }
                query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                if (i < (Porciones.size()-1)) {
                    query += " UNION ";
                }
            }
            //SI NO HAY NINGUNA PORCION FILTRADA HACER ESTO
            if (Porciones.size() == 0) {
                //CICLO QUE AGREGA TODAS LAS PORCIONES EN UNA LISTA LOCAL
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }
                query += "SELECT codigo_porcion"; //QUERY CON TODAS LAS PORCIONES
                for (int i = 0; i < Vigencias.size(); i++) {
                    query += ", COUNT(*) FILTER(WHERE (lectura_act != '' AND lectura_ant != '') AND (lectura_act - lectura_ant < 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + "'";
                }
                query += " FROM LECTURAS GROUP BY codigo_porcion";
            }

            //LISTAR VALORES EN UNA LISTA CON LISTAS DE VIGENCIAS
            List<List<String>> resultXvig = new ArrayList<List<String>>();
            for(int i = 0; i < Vigencias.size(); i++){
                resultXvig.add(new ArrayList<String>());
            }

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                //VIGENCIAS
                for (int j = 0; j < Vigencias.size(); j++) {
                    String anomXvig = rs.getString(Vigencias.get(j));
                    resultXvig.get(j).add(anomXvig);
                }
            }
            con.close(); //CERRAR CONEXION

            File fileCONSUMOS_NEGATIVOS = new File("files\\CONSUMOS_NEGATIVOS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeCONSUMOS_NEGATIVOS = new PrintWriter(fileCONSUMOS_NEGATIVOS); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "PORCION,"; //ESTRUCTURA PRIMERA FILA
            //ESCRIBIR VIGENCIAS EN LA ESTRUCTURA - PRIMERA FILA
            for (int j = 0; j < Vigencias.size(); j++) {
                estructura += ("VIG" + Vigencias.get(j));
                if (j < (Vigencias.size()-1)) {
                    estructura += ",";
                }
            }
            writeCONSUMOS_NEGATIVOS.println(estructura);
            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                writeCONSUMOS_NEGATIVOS.print(porcionesLocal.get(i));
                for (int j = 0; j < Vigencias.size(); j++) {
                    writeCONSUMOS_NEGATIVOS.print("," + resultXvig.get(j).get(i));
                }
                writeCONSUMOS_NEGATIVOS.println();
            }
            writeCONSUMOS_NEGATIVOS.println("TOTAL");
            writeCONSUMOS_NEGATIVOS.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wbCONSUMOS_NEGATIVOS = new Workbook("files\\CONSUMOS_NEGATIVOS.csv"); //NUEVO LIBRO
            Worksheet wsCONSUMOS_NEGATIVOS = wbCONSUMOS_NEGATIVOS.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (wsCONSUMOS_NEGATIVOS.getCells().getCell(0,wsCONSUMOS_NEGATIVOS.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = wsCONSUMOS_NEGATIVOS.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA ANOM

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wbCONSUMOS_NEGATIVOS.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = wsCONSUMOS_NEGATIVOS.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsCONSUMOS_NEGATIVOS.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = wsCONSUMOS_NEGATIVOS.getCells().createRange("A1:A"+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR A LA DERECHA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsCONSUMOS_NEGATIVOS.getCells().createRange("B1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);

            Cell cell;
            for (int j = 0; j < Vigencias.size(); j++) {
                if (j < Vigencias.size()) {
                    String cellChar = (wsCONSUMOS_NEGATIVOS.getCells().getCell(0,(j+1)).getName()).replaceAll("1","");
                    cell = wsCONSUMOS_NEGATIVOS.getCells().get(cellChar + (porcionesLocal.size()+2));
                    cell.setFormula("=SUM(" + cellChar + "2:" + cellChar + (porcionesLocal.size()+1) + ")");
                }
            }
            //GRAFICAR
            //CREAR GRAFICA 'TOTAL ANOMALIAS SIN 18 Y 28' Y POSICIONARLA
            int idx1 = wsCONSUMOS_NEGATIVOS.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), 0, ((porcionesLocal.size()+3)+22), (Vigencias.size()+1));
            Chart ch1 = wsCONSUMOS_NEGATIVOS.getCharts().get(idx1);
            ch1.getTitle().setText("TOTAL CONSUMOS NEGATIVOS LECTURA " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
            ch1.getNSeries().add("A"+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
            ch1.getNSeries().setCategoryData("=B1:" + lastCell + "1"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
            ch1.getNSeries().get(0).setName("=A"+(porcionesLocal.size()+2)); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
            ch1.getNSeries().get(0).setValues("=B"+(porcionesLocal.size()+2)+":" + lastCell + +(porcionesLocal.size()+2)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
            ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
            ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFIC

            wbCONSUMOS_NEGATIVOS.save("files\\CONSUMOS_NEGATIVOS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileCONSUMOS_NEGATIVOS.delete(); //ELIMINAR ARCHIVO DE CONSUMOS NEGATIVOS.csv

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> ANOMALIAS
    public void infoANOMALIAS() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y PROCEDER AL ULTIMO METODO
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR PORCIONES
            String CODPOR = "(";
            for (int j = 0; j < Porciones.size(); j++) {
                CODPOR += "codigo_porcion = '" + Porciones.get(j) + "'";
                if (j < (Porciones.size()-1)) {
                    CODPOR += " OR ";
                }
            }
            CODPOR += ") AND ";
            //SI NO SE FILTRO PORCIONES VACIAR EL STRING CON EL QUERY
            if (Porciones.size() == 0) {
                CODPOR = "";
            }

            //LISTAR OPERARIOS
            String CODOPE = "(";
            String nameOperarios = "";

            //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
            for (int j = 0; j < Operarios.size(); j++) {
                CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                nameOperarios += Operarios.get(j);
                if (j < (Operarios.size()-1)) {
                    CODOPE += " OR ";
                    nameOperarios += "-";
                }
            }
            nameOperarios = "\nOPERARIOS (" + nameOperarios + ")";
            CODOPE += ") AND ";
            //SI NO HAY OPERARIOS FILTRADOS, VACIAR LOS STRINGS Y LISTAR TODAS LAS PORCIONES EN LA LISTA LOCAL
            if (Operarios.size() == 0) {
                CODOPE = "";
                nameOperarios = "";
            }

            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            for (int j = 0; j < Vigencias.size(); j++) {
                query += "COUNT(anomalia_1) FILTER(WHERE "+ CODPOR + CODOPE +"(vigencia = '"+Vigencias.get(j)+"')) AS '"+Vigencias.get(j)+"'";
                if (j < (Vigencias.size()-1)) {
                    query += ", ";
                }
            }

            List<String> ANOMALIAS = new ArrayList<String>(); //LISTAR ANOMALIAS
            List<String> DESCRIPCION = new ArrayList<String>(); //LISTAR DESCRIPCION
            //LISTAR VALORES EN UNA LISTA CON LISTAS DE VIGENCIAS
            List<List<String>> resultXvig = new ArrayList<List<String>>();
            for(int i = 0; i < Vigencias.size(); i++){
                resultXvig.add(new ArrayList<String>());
            }

            //CONSULTA - QUERY
            PreparedStatement ps = con.prepareStatement("SELECT ANOMALIAS.ANOM, ANOMALIAS.DESCRIPCION, "+query+" FROM ANOMALIAS INNER JOIN LECTURAS ON LECTURAS.anomalia_1=ANOMALIAS.ANOM GROUP BY anomalia_1");
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                //ANOMALIAS
                String ANOM = rs.getString("ANOM");
                ANOMALIAS.add(ANOM);
                //DESCRIPCION
                String DESC = rs.getString("DESCRIPCION");
                DESCRIPCION.add(DESC);
                //VIGENCIAS
                for (int j = 0; j < Vigencias.size(); j++) {
                    String anomXvig = rs.getString(Vigencias.get(j));
                    resultXvig.get(j).add(anomXvig);
                }
            }
            con.close(); //CERRAR CONEXION

            File fileANOMALIAS = new File("files\\ANOMALIAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeANOMALIAS = new PrintWriter(fileANOMALIAS); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = "ANOM,DESCRIPCION,"; //ESTRUCTURA PRIMERA FILA
            //ESCRIBIR VIGENCIAS EN LA ESTRUCTURA - PRIMERA FILA
            for (int j = 0; j < Vigencias.size(); j++) {
                estructura += ("VIG" + Vigencias.get(j));
                if (j < (Vigencias.size()-1)) {
                    estructura += ",";
                }
            }
            writeANOMALIAS.println(estructura);
            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int j = 0; j < 26; j++) {
                writeANOMALIAS.print(ANOMALIAS.get(j) + "," + DESCRIPCION.get(j));
                for (int i = 0; i < Vigencias.size(); i++) {
                    writeANOMALIAS.print("," + resultXvig.get(i).get(j));
                }
                writeANOMALIAS.println();
            }
            writeANOMALIAS.println(",TOTAL");
            writeANOMALIAS.println(",TOTAL SIN ANOM 18 Y 28");
            writeANOMALIAS.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wbANOMALIAS = new Workbook("files\\ANOMALIAS.csv"); //NUEVO LIBRO
            Worksheet wsANOMALIAS = wbANOMALIAS.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (wsANOMALIAS.getCells().getCell(0,wsANOMALIAS.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            StyleFlag flagTOTAL = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = wsANOMALIAS.getCells();
            cells.setColumnWidth(0, 5.86); //COLUMNA ANOM
            cells.setColumnWidth(1, 27); //COLUMNA DESCRIPCION

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wbANOMALIAS.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = wsANOMALIAS.getCells().createRange("A1:"+lastCell+"27"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIAS.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS CELDAS 25 Y 26 DE DESCRIPCION
            style.setForegroundColor(com.aspose.cells.Color.getYellow()); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIAS.getCells().createRange("B25:B26"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS ANOM Y DESCRIPCION
            style.setHorizontalAlignment(TextAlignmentType.LEFT); //ALINEAR A LA IZQUIERDA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsANOMALIAS.getCells().createRange("A1:B27"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.RIGHT); //ALINEAR A LA DERECHA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsANOMALIAS.getCells().createRange("C1:"+lastCell+"27"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(9);
            //APLICAR DISEÑO Y  FORMULA DE SUMAR PARA TOTALIZAR ANOMALIAS x VIGENCIA
            style = wbANOMALIAS.createStyle();
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flagTOTAL.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flagTOTAL.setFontSize(true); //GUARDAR TAMAÑO
            range = wsANOMALIAS.getCells().createRange("B28:"+lastCell+"29"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flagTOTAL); //APLICAR DISEÑO AL RANGO DE CELDAS
            Cell cell;
            for (int j = 0; j < Vigencias.size(); j++) {
                if (j < Vigencias.size()) {
                    String cellChar = (wsANOMALIAS.getCells().getCell(0,(j+2)).getName()).replaceAll("1","");
                    cell = wsANOMALIAS.getCells().get(cellChar + "28");
                    cell.setFormula("=SUM(" + cellChar + "2:" + cellChar + "27)");
                    cell = wsANOMALIAS.getCells().get(cellChar + "29");
                    cell.setFormula("=SUM("+cellChar+"2:"+cellChar+"15)+SUM("+cellChar+"17:"+cellChar+"24)+SUM("+cellChar+"26:"+cellChar+"27)");
                }
            }
            //GRAFICAR
            //CREAR GRAFICA 'TOTAL ANOMALIAS SIN 18 Y 28' Y POSICIONARLA
            int idx1 = wsANOMALIAS.getCharts().add(ChartType.LINE, 30, 0, 47, (Vigencias.size()+2));
            Chart ch1 = wsANOMALIAS.getCharts().get(idx1);
            ch1.getTitle().setText("TOTAL ANOMALIAS SIN 18 Y 28 " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
            ch1.getNSeries().add("B29", true); //AGREGA LA SERIE
            ch1.getNSeries().setCategoryData("=C1:" + lastCell + "1"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
            ch1.getNSeries().get(0).setName("=B29"); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
            ch1.getNSeries().get(0).setValues("=C29:" + lastCell + "29"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
            ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
            ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
            //CREAR GRAFICA 'TOTAL ANOMALIAS 18' Y POSICIONARLA
            int idx2 = wsANOMALIAS.getCharts().add(ChartType.LINE, 47, 0, 64, (Vigencias.size()+2));
            Chart ch2 = wsANOMALIAS.getCharts().get(idx2);
            ch2.getTitle().setText("TOTAL ANOMALIAS 18 " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch2.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch2.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            ch2.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
            ch2.getNSeries().add("B16", true); //AGREGA LA SERIE
            ch2.getNSeries().setCategoryData("=C1:" + lastCell + "1"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
            ch2.getNSeries().get(0).setName("=B16"); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
            ch2.getNSeries().get(0).setValues("=C16:" + lastCell + "16"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            ch2.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
            ch2.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
            ch2.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
            //CREAR GRAFICA 'TOTAL ANOMALIAS 28' Y POSICIONARLA
            int idx3 = wsANOMALIAS.getCharts().add(ChartType.LINE, 64, 0, 81, (Vigencias.size()+2));
            Chart ch3 = wsANOMALIAS.getCharts().get(idx3);
            ch3.getTitle().setText("TOTAL ANOMALIAS 28 " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch3.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch3.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            ch3.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
            ch3.getNSeries().add("B25", true); //AGREGA LA SERIE
            ch3.getNSeries().setCategoryData("=C1:" + lastCell + "1"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
            ch3.getNSeries().get(0).setName("=B25"); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
            ch3.getNSeries().get(0).setValues("=C25:" + lastCell + "25"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            ch3.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
            ch3.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
            ch3.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

            wbANOMALIAS.save("files\\ANOMALIAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS EN UN ARCHIVO EXCEL
            fileANOMALIAS.delete(); //ELIMINAR ARCHIVO DE ANOMALIAS.csv

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> ANOMALIASxPORCION
    public void infoANOMALIASxPORCION() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y FINALIZAR LA PANTALLA DE CARGA
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = " (";
            String nameOperarios = "";
            for (int j = 0; j < Operarios.size(); j++) {
                CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                nameOperarios += Operarios.get(j);
                if (j < (Operarios.size() - 1)) {
                    CODOPE += " OR ";
                    nameOperarios += "-";
                }
            }
            nameOperarios = "\nOPERARIOS (" + nameOperarios + ")";
            CODOPE += ") AND";

            //SI NO SE FILTRO OPERARIOS VACIAR EL STRING CON EL QUERY
            if (Operarios.size() == 0) {
                CODOPE = "";
                nameOperarios = "";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA REALIZAR EL CICLO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i));
                query += "SELECT codigo_porcion";
                for (int j = 0; j < Vigencias.size(); j++) {
                    query += ", COUNT (*) FILTER (WHERE (codigo_porcion = '" + Porciones.get(i) + "') AND" + CODOPE + " (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":TOTAL', COUNT (*) FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND (codigo_porcion = '" + Porciones.get(i) + "') AND" + CODOPE + " (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":FILTRADO', printf(\"%.6f\",(COUNT() FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODOPE + " (vigencia = '" + Vigencias.get(j) + "'))*1.0/COUNT() FILTER(WHERE" + CODOPE + " (vigencia = '" + Vigencias.get(j) + "')))) AS '" + Vigencias.get(j) + ":PORCENTAJE'";
                }
                query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                if (i < (Porciones.size() - 1)) {
                    query += " UNION ";
                }
            }

            //SI NO HAY NINGUNA PORCION FILTRADA HACER ESTO
            if (Porciones.size() == 0) {
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }
                query += "SELECT codigo_porcion";
                for (int i = 0; i < Vigencias.size(); i++) {
                    query += ", COUNT (*) FILTER(WHERE" + CODOPE + " (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":TOTAL', COUNT (*) FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODOPE + " (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":FILTRADO', printf(\"%.6f\",(COUNT() FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODOPE + " (vigencia = '" + Vigencias.get(i) + "'))*1.0/COUNT() FILTER(WHERE" + CODOPE + " (vigencia = '" + Vigencias.get(i) + "')))) AS '" + Vigencias.get(i) + ":PORCENTAJE'";
                }
                query += " FROM LECTURAS GROUP BY codigo_porcion";
            }

            //LISTAR VALORES EN UNA LISTA CON LISTAS DE VIGENCIAS
            List<List<String>> resultXvig = new ArrayList<List<String>>();
            for (int i = 0; i < Vigencias.size(); i++) {
                resultXvig.add(new ArrayList<String>());
            }

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                //VIGENCIAS
                for (int i = 0; i < Vigencias.size(); i++) {
                    String totalXvig = rs.getString(Vigencias.get(i) + ":TOTAL");
                    String filtradoXvig = rs.getString(Vigencias.get(i) + ":FILTRADO");
                    String porcentajeXvig = rs.getString(Vigencias.get(i) + ":PORCENTAJE");
                    porcentajeXvig = "\"" + porcentajeXvig.replace(".", ",") + "\"";
                    resultXvig.get(i).add(totalXvig + "," + filtradoXvig + "," + porcentajeXvig);
                }
            }

            con.close(); //CERRAR CONEXION

            File fileANOMALIASxPORCION = new File("files\\ANOMALIASxPORCION.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeANOMALIASxPORCION = new PrintWriter(fileANOMALIASxPORCION); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = ","; //ESTRUCTURA PRIMERA FILA
            //CICLO QUE ESCRIBE LAS VIGENCIAS EN LA ESTRUCTURA Y SEPARANDO DOS CELDAS POR LAS SUBCOLUMNAS DE CADA VIGENCIA
            for (int i = 0; i < Vigencias.size(); i++) {
                estructura += ("VIG "+Vigencias.get(i));
                if (i < (Vigencias.size() - 1)) {
                    estructura += ",,,";
                }
            }
            writeANOMALIASxPORCION.println(estructura);
            //ESTRUCTURA SEGUNDA FILA PORCION -> TOTAL, FILTRADO, % DE CADA VIGENCIA
            estructura = "PORCION,";
            for (int i = 0; i < Vigencias.size(); i++) {
                estructura += "TOTAL,FILTRADO,%";
                if (i < (Vigencias.size() - 1)) {
                    estructura += ",";
                }
            }
            writeANOMALIASxPORCION.println(estructura);

            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                writeANOMALIASxPORCION.print(porcionesLocal.get(i));
                for (int j = 0; j < Vigencias.size(); j++) {
                    writeANOMALIASxPORCION.print("," + resultXvig.get(j).get(i));
                }
                writeANOMALIASxPORCION.println();
            }
            writeANOMALIASxPORCION.close();
            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wbANOMALIASxPORCION = new Workbook("files\\ANOMALIASxPORCION.csv"); //NUEVO LIBRO
            Worksheet wsANOMALIASxPORCION = wbANOMALIASxPORCION.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (wsANOMALIASxPORCION.getCells().getCell(0,wsANOMALIASxPORCION.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = wsANOMALIASxPORCION.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA ANOM

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wbANOMALIASxPORCION.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = wsANOMALIASxPORCION.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //GRAFICAR
            //CREAR GRAFICA 'TOTAL ANOMALIAS SIN 18 Y 28' Y POSICIONARLA
            int idx1 = wsANOMALIASxPORCION.getCharts().add(ChartType.COLUMN, (porcionesLocal.size()+3), 0, ((porcionesLocal.size()+3)+22), (Vigencias.size()*3)+1);
            Chart ch1 = wsANOMALIASxPORCION.getCharts().get(idx1);
            ch1.getTitle().setText("TOTAL ANOMALIAS x PORCION LECTURA " + nameOperarios); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            //AGREGAR PORCENTAJES Y GRAFICAR
            int column = 0;
            for (int i = 0; i < Vigencias.size(); i++) {
                column += 3; //POR CADA VIGENCIA TOMAR LA TERCERA SUBCOLUMNA
                range = wsANOMALIASxPORCION.getCells().createRange(wsANOMALIASxPORCION.getCells().getCell(2,column).getName() + ":" + wsANOMALIASxPORCION.getCells().getCell((porcionesLocal.size()+1),column).getName()); //TOMAR RANGO DE CELDAS
                style.setNumber(10); //CONVERTIR NUMERO DE CELDA EN PORCENTAJE
                range.setStyle(style); //APLICAR DISEÑO AL RANGO DE CELDAS
                cells.merge(0, column-2, 1, 3); //COMBINAR Y CENTRAR 3 COLUMNAS POR CADA VIGENCIA
                //GRAFICAR DE CADA PORCENTAJE OBTENIDO
                ch1.getNSeries().add(Vigencias.get(i), false); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData("=A3:A" + (porcionesLocal.size()+2)); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                ch1.getNSeries().get(i).setName("VIG " + Vigencias.get(i) ); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                ch1.getNSeries().get(i).setValues(range.getRefersTo()); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            }

            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIASxPORCION.getCells().createRange("B2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = wsANOMALIASxPORCION.getCells().createRange("A2:A"+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142));
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIASxPORCION.getCells().createRange("B1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS

            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR A LA DERECHA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsANOMALIASxPORCION.getCells().createRange("B1:"+lastCell+(porcionesLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);

            wbANOMALIASxPORCION.save("files\\ANOMALIASxPORCION.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileANOMALIASxPORCION.delete(); //ELIMINAR ARCHIVO DE ANOMALIASxPORCION.csv
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        valFINISH += 1;
        if (valINIT == valFINISH) {
            INFORME();
        }
    }

    //METODO informe -> ANOMALIASxOPERARIO
    public void infoANOMALIASxOPERARIO() {
        valINIT += 1; //INICIA METODO SUMA valINIT PARA VALIDAR AL FINAL DEL METODO SI TODOS LOS METODOS QUE INICIARON AL MISMO TIEMPO TERMINARON Y FINALIZAR LA PANTALLA DE CARGA
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR PORCIONES
            String CODPOR = " (";
            String namePorciones = "";
            for (int j = 0; j < Porciones.size(); j++) {
                CODPOR += "codigo_porcion = '" + Porciones.get(j) + "'";
                namePorciones += Porciones.get(j);
                if (j < (Porciones.size() - 1)) {
                    CODPOR += " OR ";
                    namePorciones += "-";
                }
            }
            namePorciones = "\nPORCIONES (" + namePorciones + ")";
            CODPOR += ") AND";

            //SI NO SE FILTRO PORCIONES VACIAR EL STRING CON EL QUERY
            if (Porciones.size() == 0) {
                CODPOR = "";
                namePorciones = "";
            }

            //LISTAR PORCIONES
            ArrayList<String> operariosLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUN OPERARIO ESTA FILTRADO REALIZAR EL CICLO
            for (int i = 0; i < Operarios.size(); i++) {
                operariosLocal.add(Operarios.get(i));
                query += "SELECT codigo_operario";
                for (int j = 0; j < Vigencias.size(); j++) {
                    query += ", COUNT (*) FILTER (WHERE (codigo_operario = '" + Operarios.get(i) + "') AND" + CODPOR + " (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":TOTAL', COUNT (*) FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND (codigo_operario = '" + Operarios.get(i) + "') AND" + CODPOR + " (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":FILTRADO', printf(\"%.6f\",(COUNT() FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODPOR + " (vigencia = '" + Vigencias.get(j) + "'))*1.0/COUNT() FILTER(WHERE" + CODPOR + " (vigencia = '" + Vigencias.get(j) + "')))) AS '" + Vigencias.get(j) + ":PORCENTAJE'";
                }
                query += " FROM LECTURAS WHERE (codigo_operario = '" + Operarios.get(i) + "')";
                if (i < (Operarios.size() - 1)) {
                    query += " UNION ";
                }
            }

            //SI NO HAY NINGUN OPERARIO FILTRADA HACER ESTO
            if (Operarios.size() == 0) {
                for (int i = 0; i < CHBX_CODOPE.length; i++) {
                    operariosLocal.add(CHBX_CODOPE[i].getText());
                }
                query += "SELECT codigo_operario";
                for (int i = 0; i < Vigencias.size(); i++) {
                    query += ", COUNT (*) FILTER(WHERE" + CODPOR + " (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":TOTAL', COUNT (*) FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODPOR + " (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":FILTRADO', printf(\"%.6f\",(COUNT() FILTER(WHERE (anomalia_1 != '') AND (anomalia_1 = 9 OR anomalia_1 = 16 OR anomalia_1 = 17 OR anomalia_1 = 19 OR anomalia_1 = 20) AND " + CODPOR + " (vigencia = '" + Vigencias.get(i) + "'))*1.0/COUNT() FILTER(WHERE" + CODPOR + " (vigencia = '" + Vigencias.get(i) + "')))) AS '" + Vigencias.get(i) + ":PORCENTAJE'";
                }
                query += " FROM LECTURAS GROUP BY codigo_operario";
            }

            //LISTAR VALORES EN UNA LISTA CON LISTAS DE VIGENCIAS
            List<List<String>> resultXvig = new ArrayList<List<String>>();
            for (int i = 0; i < Vigencias.size(); i++) {
                resultXvig.add(new ArrayList<String>());
            }

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                //VIGENCIAS
                for (int i = 0; i < Vigencias.size(); i++) {
                    String totalXvig = rs.getString(Vigencias.get(i) + ":TOTAL");
                    String filtradoXvig = rs.getString(Vigencias.get(i) + ":FILTRADO");
                    String porcentajeXvig = rs.getString(Vigencias.get(i) + ":PORCENTAJE");
                    porcentajeXvig = "\"" + porcentajeXvig.replace(".", ",") + "\"";
                    resultXvig.get(i).add(totalXvig + "," + filtradoXvig + "," + porcentajeXvig);
                }
            }

            con.close(); //CERRAR CONEXION

            File fileANOMALIASxOPERARIO = new File("files\\ANOMALIASxOPERARIO.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter writeANOMALIASxOPERARIO = new PrintWriter(fileANOMALIASxOPERARIO); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO

            String estructura = ","; //ESTRUCTURA PRIMERA FILA
            //CICLO QUE ESCRIBE LAS VIGENCIAS EN LA ESTRUCTURA Y SEPARANDO DOS CELDAS POR LAS SUBCOLUMNAS DE CADA VIGENCIA
            for (int i = 0; i < Vigencias.size(); i++) {
                estructura += ("VIG "+Vigencias.get(i));
                if (i < (Vigencias.size() - 1)) {
                    estructura += ",,,";
                }
            }
            writeANOMALIASxOPERARIO.println(estructura);
            //ESTRUCTURA SEGUNDA FILA PORCION -> TOTAL, FILTRADO, % DE CADA VIGENCIA
            estructura = "OPERARIO,";
            for (int i = 0; i < Vigencias.size(); i++) {
                estructura += "TOTAL,FILTRADO,%";
                if (i < (Vigencias.size() - 1)) {
                    estructura += ",";
                }
            }
            writeANOMALIASxOPERARIO.println(estructura);

            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < operariosLocal.size(); i++) {
                writeANOMALIASxOPERARIO.print(operariosLocal.get(i));
                for (int j = 0; j < Vigencias.size(); j++) {
                    writeANOMALIASxOPERARIO.print("," + resultXvig.get(j).get(i));
                }
                writeANOMALIASxOPERARIO.println();
            }
            writeANOMALIASxOPERARIO.close();
            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wbANOMALIASxOPERARIO = new Workbook("files\\ANOMALIASxOPERARIO.csv"); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Worksheet wsANOMALIASxOPERARIO = wbANOMALIASxOPERARIO.getWorksheets().get(0); //NUEVA HOJA DE ANOMALIAS PARA EL LIBRO DE ANOMALIAS

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (wsANOMALIASxOPERARIO.getCells().getCell(0,wsANOMALIASxOPERARIO.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = wsANOMALIASxOPERARIO.getCells();
            cells.setColumnWidth(0, 9); //COLUMNA OPERARIOS

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wbANOMALIASxOPERARIO.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = wsANOMALIASxOPERARIO.getCells().createRange("A1:"+lastCell+(operariosLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //GRAFICAR
            //CREAR GRAFICA 'TOTAL ANOMALIAS SIN 18 Y 28' Y POSICIONARLA
            int idx1 = wsANOMALIASxOPERARIO.getCharts().add(ChartType.COLUMN, (operariosLocal.size()+3), 0, ((operariosLocal.size()+3)+22), (Vigencias.size()*3)+1);
            Chart ch1 = wsANOMALIASxOPERARIO.getCharts().get(idx1);
            ch1.getTitle().setText("TOTAL ANOMALIAS x OPERARIO LECTURA " + namePorciones); //ASIGNARLE UN NOMBRE A LA GRAFICA
            ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
            ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
            //AGREGAR PORCENTAJES Y GRAFICAR
            int column = 0;
            for (int i = 0; i < Vigencias.size(); i++) {
                column += 3; //POR CADA VIGENCIA TOMAR LA TERCERA SUBCOLUMNA
                range = wsANOMALIASxOPERARIO.getCells().createRange(wsANOMALIASxOPERARIO.getCells().getCell(2,column).getName() + ":" + wsANOMALIASxOPERARIO.getCells().getCell((operariosLocal.size()+1),column).getName()); //TOMAR RANGO DE CELDAS
                style.setNumber(10); //CONVERTIR NUMERO DE CELDA EN PORCENTAJE
                range.setStyle(style); //APLICAR DISEÑO AL RANGO DE CELDAS
                cells.merge(0, column-2, 1, 3); //COMBINAR Y CENTRAR 3 COLUMNAS POR CADA VIGENCIA
                //GRAFICAR DE CADA PORCENTAJE OBTENIDO
                ch1.getNSeries().add(Vigencias.get(i), false); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData("=A3:A" + (operariosLocal.size()+2)); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                ch1.getNSeries().get(i).setName("VIG " + Vigencias.get(i) ); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                ch1.getNSeries().get(i).setValues(range.getRefersTo()); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
            }

            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIASxOPERARIO.getCells().createRange("B2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = wsANOMALIASxOPERARIO.getCells().createRange("A2:A"+(operariosLocal.size()+2)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142));
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = wsANOMALIASxOPERARIO.getCells().createRange("B1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS

            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR A LA DERECHA EN HORIZONTAL
            style.setVerticalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN VERTICAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = wsANOMALIASxOPERARIO.getCells().createRange("B1:"+lastCell+(operariosLocal.size()+2)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);

            wbANOMALIASxOPERARIO.save("files\\ANOMALIASxOPERARIO.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            fileANOMALIASxOPERARIO.delete(); //ELIMINAR ARCHIVO DE ANOMALIASxPORCION.csv
        } catch (Exception ex) {
            ex.printStackTrace();
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

            //LISTAR CODIGO PORCION & LEIDO, NO LEIDO Y TOTAL
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
                estructura += "LEIDO,NO LEIDO,TOTAL";
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
            Workbook wbINFORME = new Workbook(); //NUEVO LIBRO
            //SELECCIONAR LOS LIBROS CON LAS TABLAS
            File fileEXCEL_CONSUMO_0 = new File("files\\CONSUMO_0.xlsx");
            File fileEXCEL_CONSUMOS_NEGATIVOS = new File("files\\CONSUMOS_NEGATIVOS.xlsx");
            File fileEXCEL_ANOMALIAS = new File("files\\ANOMALIAS.xlsx");
            File fileEXCEL_ANOMALIASxPORCION = new File("files\\ANOMALIASxPORCION.xlsx");
            File fileEXCEL_ANOMALIASxOPERARIO = new File("files\\ANOMALIASxOPERARIO.xlsx");
            Workbook wbCONSUMO_0 = new Workbook(fileEXCEL_CONSUMO_0.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook wbCONSUMOS_NEGATIVOS = new Workbook(fileEXCEL_CONSUMOS_NEGATIVOS.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook wbANOMALIAS = new Workbook(fileEXCEL_ANOMALIAS.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook ANOMALIASxPORCION = new Workbook(fileEXCEL_ANOMALIASxPORCION.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            Workbook ANOMALIASxOPERARIO = new Workbook(fileEXCEL_ANOMALIASxOPERARIO.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
            //COMBINAR HOJAS EN EL INFORME
            wbINFORME.combine(wbCONSUMO_0);
            wbINFORME.combine(wbCONSUMOS_NEGATIVOS);
            wbINFORME.combine(wbANOMALIAS);
            wbINFORME.combine(ANOMALIASxPORCION);
            wbINFORME.combine(ANOMALIASxOPERARIO);
            wbINFORME.getWorksheets().removeAt(0); //ELIMINAR LA PRIMERA HOJA VACIA DEL LIBRO
            wbINFORME.save("files\\INFORME.xlsx");
            Thread.sleep(2*1000);
            //ELIMINAR LIBROS COPIADOS
            fileEXCEL_CONSUMO_0.delete();
            fileEXCEL_CONSUMOS_NEGATIVOS.delete();
            fileEXCEL_ANOMALIAS.delete();
            fileEXCEL_ANOMALIASxPORCION.delete();
            fileEXCEL_ANOMALIASxOPERARIO.delete();

            dialog.dispose(); //CERRAR LOADING
            JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE EL INFORME", "",JOptionPane.INFORMATION_MESSAGE);
            File ARCHIVOS = new File("files");
            Runtime.getRuntime().exec("cmd /c start " + ARCHIVOS.getAbsolutePath() + " && exit");
        } catch (Exception ex) {
            ex.printStackTrace();
            dialog.dispose(); //CERRAR LOADING
            JOptionPane.showMessageDialog(null, "ERROR INESPERADO. INTENTE NUEVAMENTE", "",JOptionPane.INFORMATION_MESSAGE);
        }
    }

    //METODO MAIN
    public static void main(String[] args) {
        new PROGRAMA();
    }

}

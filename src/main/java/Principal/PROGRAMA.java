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
import java.awt.Label;
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

    //--------VALIDAR INICIO Y FIN DE PROCESOS--------
    int INITprogram;
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
            panelLOAD.add(new JLabel("          CARGANDO PROGRAMA"), BorderLayout.PAGE_START); //AÑADIR UN LABEL AL INICIO DEL PANEL
        }

        panelLOAD.add(pbCargar, BorderLayout.CENTER); //AÑADIR BARRA DE PROGRESO EN EL CENTRO DEL PANEL

        if (INITprogram == 0) {
            JButton btnEXIT = new JButton("x");
            btnEXIT.setPreferredSize(new Dimension(50,15));
            panelLOAD.add(btnEXIT, BorderLayout.LINE_END); //AÑADIR UN BOTON PARA CANCELAR EL PROGRAMA CUANDO EMPIECE A CARGAR
            //ACCION BOTON EXIT
            btnEXIT.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    System.exit(0);
                }
            });
        }

        panelLOAD.setBackground(Color.CYAN); //ASIGNAR COLOR AZUL AL PANEL
        dialog = new JDialog(frameLOAD, true);

        dialog.setUndecorated(true);
        dialog.getContentPane().add(panelLOAD);
        dialog.pack();
        dialog.setLocationRelativeTo(null);
        dialog.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        dialog.setVisible(true);
    }

    //METODO INIT
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

            JCheckBox CHBX_SELECTALL = new JCheckBox("Seleccionar todo");
            jpCHECK_CODOPE.add(CHBX_SELECTALL);

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

            for (int j = 0; j < Operarios.size(); j++) {
                CHBX_CODOPE[j].addActionListener(new ActionListener() {
                    @Override
                    public void actionPerformed(ActionEvent e) {
                        if (CHBX_SELECTALL.isSelected()) {
                            CHBX_SELECTALL.setSelected(false);
                        } else {
                            Boolean bol = false;
                            for (int j = 0; j < Operarios.size(); j++) {
                                if (CHBX_CODOPE[j].isSelected()) {
                                    bol = true;
                                } else {
                                    bol = false;
                                    break;
                                }
                            }
                            if (bol.equals(true)) {
                                CHBX_SELECTALL.setSelected(true);
                            }
                        }
                    }
                });
            }

            CHBX_SELECTALL.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    if (CHBX_SELECTALL.isSelected()) {
                        for (int j = 0; j < Operarios.size(); j++) {
                            CHBX_CODOPE[j].setSelected(true);
                        }
                    } else {
                        for (int j = 0; j < Operarios.size(); j++) {
                            CHBX_CODOPE[j].setSelected(false);
                        }
                    }
                }
            });

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

                        if (codigo_porcion.charAt(0) == 'W' || codigo_porcion.charAt(0) == 'X' || codigo_porcion.charAt(0) == 'Z') {
                            if (vigencia.charAt(5) == '0') {
                                codigo_porcion += "-1";
                            } else {
                                codigo_porcion += "-2";
                                StringBuilder nuevaVIGENCIA = new StringBuilder(vigencia);
                                nuevaVIGENCIA.setCharAt(5,'0');
                                vigencia = nuevaVIGENCIA.toString();
                            }
                        }
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

            Collections.sort(Vigencias); //ORDENAR VIGENCIAS DE MENOR A MAYOR

            //INICIAR METODOS
            new Thread(() -> infoLECTURAS()).start();
            new Thread(() -> infoCONSUMO_0()).start();
            new Thread(() -> infoCONSUMOS_NEGATIVOS()).start();
            new Thread(() -> infoANOMALIAS()).start();
            //new Thread(() -> infoANOMALIASxPORCION()).start();
            //new Thread(() -> infoANOMALIASxOPERARIO()).start();

        }
    }

    //METODO informe -> LECTURAS
    public void infoLECTURAS() {
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = "";
            //SI LA CANTIDAD DE OPERARIOS FILTRADOS ES DIFERENTE A 0 Y A LA CANTIDAD TOTAL EXISTENTES HACER ESTO
            if (Operarios.size() != 0 && Operarios.size() != CHBX_CODOPE.length) {
                CODOPE = " AND (";
                //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
                for (int j = 0; j < Operarios.size(); j++) {
                    CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                    if (j < (Operarios.size() - 1)) {
                        CODOPE += " OR ";
                    }
                }
                CODOPE += ")";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA HACER ESTO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i)); //AGREGAR PORCIONES FILTRADAS A LA LISTA LOCAL
                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":NOLEIDO', COUNT(*) FILTER(WHERE (vigencia = '" + Vigencias.get(j) + "')" + CODOPE + ") AS '" + Vigencias.get(j) + ":TOTAL'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int j = 0; j < Operarios.size(); j++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(j) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int l = 0; l < Vigencias.size(); l++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '') AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_operario = '" + Operarios.get(j) + "') AND (vigencia = '" + Vigencias.get(l) + "')) AS '" + Vigencias.get(l) + ":" + Operarios.get(j) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '') AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_operario = '" + Operarios.get(j) + "') AND (vigencia = '" + Vigencias.get(l) + "')) AS '" + Vigencias.get(l) + ":" + Operarios.get(j) + ":NOLEIDO', COUNT(*) FILTER(WHERE (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_operario = '" + Operarios.get(j) + "') AND (vigencia = '" + Vigencias.get(l) + "')) AS '" + Vigencias.get(l) + ":" + Operarios.get(j) + ":TOTAL'";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int j = 0; j < Vigencias.size(); j++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += " COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '') AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '') AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":NOLEIDO', COUNT(*) FILTER(WHERE (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":TOTAL'";
                        if (j+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }
            }

            //SI NO SE FILTRO NINGUNA PORCION HACER ESTO
            if (Porciones.size() == 0) {
                //CICLO QUE AGREGA TODAS LAS PORCIONES EXISTENTES EN UNA LISTA LOCAL
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }

                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS

                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '')" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":NOLEIDO', COUNT(*) FILTER(WHERE (vigencia = '" + Vigencias.get(j) + "')" + CODOPE + ") AS '" + Vigencias.get(j) + ":TOTAL'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int i = 0; i < Operarios.size(); i++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(i) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '') AND (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":" + Operarios.get(i) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '') AND (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":" + Operarios.get(i) + ":NOLEIDO', COUNT(*) FILTER(WHERE (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":" + Operarios.get(i) + ":TOTAL'";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";

                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int i = 0; i < Vigencias.size(); i++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += "COUNT(*) FILTER(WHERE (lectura_act != '' OR anomalia_1 != '') AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":LEIDO', COUNT(*) FILTER(WHERE (lectura_act = '' AND anomalia_1 = '') AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":NOLEIDO', COUNT(*) FILTER(WHERE (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + ":TOTAL'";
                        if (i+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";
                }
            }

            List<String> resultLIST = new ArrayList(); //LISTA PARA SACAR LOS RESULTADOS DE CADA FILA

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String datosXporcion = "";

                //SI NO SE FILTRO NINGUN OPERARIO O SE FILTRO MAS DE UN OPERARIO HACER ESTO
                if (Operarios.size() == 0 || Operarios.size() > 1) {
                    //EN TOTAL = CODIGO PORCION x VIGENCIAS -> RESULTADO
                    String result = rs.getString("codigo_porcion");
                    datosXporcion += result + ",";
                    for (int i = 0; i < Vigencias.size(); i++) {
                        result = rs.getString(Vigencias.get(i) + ":LEIDO");
                        result += "," + rs.getString(Vigencias.get(i) + ":NOLEIDO");
                        result += "," + rs.getString(Vigencias.get(i) + ":TOTAL");
                        if (Operarios.size() == 0) {
                            datosXporcion += result;
                            if (i < (Vigencias.size()-1)) {
                                datosXporcion += ",";
                            }
                        } else {
                            datosXporcion += result + ",";
                        }
                    }
                }

                //CICLO POR OPERARIO = CODIGO_PORCION x VIGENCIAS -> RESULTADO
                for (int i = 0; i < Operarios.size(); i++) {
                    String result = rs.getString("codigo_porcion:" + Operarios.get(i));
                    datosXporcion += result + ",";
                    for (int j = 0; j < Vigencias.size(); j++) {
                        result = rs.getString(Vigencias.get(j) + ":" + Operarios.get(i) + ":LEIDO");
                        result += "," + rs.getString(Vigencias.get(j) + ":" + Operarios.get(i) + ":NOLEIDO");
                        result += "," + rs.getString(Vigencias.get(j) + ":" + Operarios.get(i) + ":TOTAL");
                        datosXporcion += result;
                        if (j < Vigencias.size()-1 || i < Operarios.size()-1) {
                            datosXporcion += ",";
                        }
                    }
                }
                resultLIST.add(datosXporcion);
            }
            con.close(); //CERRAR CONEXION

            File file = new File("files\\LECTURAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter write = new PrintWriter(file); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO


            String estructura = ""; //ESTRUCTURA PRIMERA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0) {
                estructura += "TODOS LOS OPERARIOS"; //TOTAL
            } else if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                estructura += "TODOS LOS OPERARIOS FILTRADOS,"; //TOTAL
                //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                for (int j = 0; j < Vigencias.size(); j++) { // +1 POR LA COLUMNA PORCION
                    estructura += ",,,";
                }
            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "OPERARIO " + Operarios.get(i);
                if (i < (Operarios.size()-1)) {
                    estructura += ",";
                }
                for (int j = 0; j < Vigencias.size(); j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",,,";
                    }
                }
            }
            write.println(estructura);
            estructura = ""; //VACIAR EL STRING

            //ESCRIBIR LAS PORCIONES Y LAS VIGENCIAS EN LA SEGUNDA FILA DE LA ESTRUCTURA
            int OyV; //ENTERO QUE SERVIRA PARA LA LONGITUD DEL CICLO
            //SI SE FILTRO SOLAMENTE 1 OPERARIO
            if (Operarios.size() == 1) {
                OyV = 1; //SOLAMENTE REPETIR EL CICLO 1 VEZ
            } else {
                OyV = Operarios.size() + 1;  //PORCIONES SELECCIONADAS + 1 DEL TOTAL
            }

            for (int i = 0; i < OyV; i++) { //CICLO POR CADA OPERARIO QUE EXISTA AGREGAR LAS VIGENCIAS EXISTENTES
                estructura += ",";
                for (int j = 0; j < Vigencias.size(); j++) {
                    estructura += ("VIG" + Vigencias.get(j));
                    if (j < (Vigencias.size()-1)) { //SI j ES MENOR AL TOTAL DE VIGENCIAS, SEPARAR LAS VIGENCIAS HASTA SER IGUAL AL TOTAL DE VIGENCIAS, ES DECIR, HASTA QUE TERMINE DE SEPARAR TODAS LAS VIGENCIAS
                        estructura += ",,,";
                    }
                }
                if (Operarios.size() > 1 && i < (Operarios.size())) { //SI SE FILTRO MAS DE UN OPERARIO Y j ES MENOR A CADA OPERARIO SEPARAR TODA LA ESTRUCTURA PARA VOLVER A REESCRIBIR LAS PORCIONES Y VIGENCIAS DE CADA OPERARIO HASTA QUE j SEA IGUAL, ES DECIR, TERMINE DE SEPARAR TODOS LOS OPERARIOS
                    estructura += ",,,";
                }
            }
            write.println(estructura);
            estructura = ""; //VACIAR EL STRING
            for (int i = 0; i < OyV; i++) { //CICLO POR CADA OPERARIO QUE EXISTA AGREGAR LAS VIGENCIAS EXISTENTES
                estructura += "PORCION,";
                for (int j = 0; j < Vigencias.size(); j++) {
                    estructura += ("LEIDO,NO LEIDO,TOTAL");
                    if (j < (Vigencias.size()-1)) { //SI j ES MENOR AL TOTAL DE VIGENCIAS, SEPARAR LAS VIGENCIAS HASTA SER IGUAL AL TOTAL DE VIGENCIAS, ES DECIR, HASTA QUE TERMINE DE SEPARAR TODAS LAS VIGENCIAS
                        estructura += ",";
                    }
                }
                if (Operarios.size() > 1 && i < (Operarios.size())) { //SI SE FILTRO MAS DE UN OPERARIO Y j ES MENOR A CADA OPERARIO SEPARAR TODA LA ESTRUCTURA PARA VOLVER A REESCRIBIR LAS PORCIONES Y VIGENCIAS DE CADA OPERARIO HASTA QUE j SEA IGUAL, ES DECIR, TERMINE DE SEPARAR TODOS LOS OPERARIOS
                    estructura += ",";
                }
            }
            write.println(estructura);
            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                write.println(resultLIST.get(i));
            }
            //AÑADIR TOTALIZADOR
            estructura = ""; //ESTRUCTURA ULTIMA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0 || Operarios.size() > 1) {
                estructura += "TOTAL"; //TOTAL
                if (Operarios.size() > 1) {
                    estructura += ",";
                }
                if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                    //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                    for (int j = 0; j < Vigencias.size(); j++) { // +1 POR LA COLUMNA PORCION
                        estructura += ",,,";
                    }
                }

            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "TOTAL";
                if (i < (Operarios.size()-1)) {
                    estructura += ",";
                }
                for (int j = 0; j < Vigencias.size(); j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",,,";
                    }
                }
            }
            write.println(estructura);
            write.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wb = new Workbook("files\\LECTURAS.csv"); //NUEVO LIBRO
            Worksheet worksheet = wb.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (worksheet.getCells().getCell(0,worksheet.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = worksheet.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA PORCION

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wb.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = worksheet.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+4)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 255, 0)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(169, 208, 142)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS SEGUNDAS FILAS Y COLUMNAS PORCION
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A3:"+lastCell+"3"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = worksheet.getCells().createRange("A3:A"+(porcionesLocal.size()+4)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN HORIZONTAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = worksheet.getCells().createRange("B2:"+lastCell+(porcionesLocal.size()+4)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);
            range = worksheet.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS

            Cell cell;
            int valor = 0;
            int columnaVIGENCIA = 0;
            String celdaVIGENCIAS = "=";
            String celdaLEIDO = "=";

            //SI NO SE FILTRO NINGUN OPERARIO O SOLO SE FILTRO 1 SOLAMENTE HACER ESTO
            if (Operarios.size() <= 1) {
                //CREAR GRAFICA 'TOTAL CONSUMO 0' Y POSICIONARLA
                cells.merge(0, 0, 1, (Vigencias.size()*3)+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                for (int i = 0; i < 1; i++) {
                    for (int j = 0; j < Vigencias.size()*3; j++) {
                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+3),valor).getName()).replaceAll(""+(porcionesLocal.size()+4),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+4));
                        cell.setFormula("=SUM(" + cellChar + "4:" + cellChar + (porcionesLocal.size()+3) + ")");
                        if (valor % 3 == 1) {
                            cells.merge(1, valor, 1, 3); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                            celdaVIGENCIAS += cellChar + "2";
                            celdaLEIDO += cellChar + (porcionesLocal.size()+4);
                            if (j < (Vigencias.size()*3)-3) {
                                celdaVIGENCIAS += ",";
                                celdaLEIDO += ",";
                            }
                        }
                    }
                    valor += 1;
                }
                int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+4), 0, ((porcionesLocal.size()+3)+16), (Vigencias.size()*3)+1);
                Chart ch1 = worksheet.getCharts().get(idx1);
                ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch1.getNSeries().add("A"+(porcionesLocal.size()+4), true); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData(celdaVIGENCIAS); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                ch1.getNSeries().get(0).setValues(celdaLEIDO); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA

                if (Operarios.size() == 0) {
                    ch1.getNSeries().get(0).setName("=\"TOTAL LECTURAS LEIDAS\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch1.getNSeries().get(0).setName("=\"TOTAL LECTURAS LEIDAS\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }
                ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
            } else { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                for (int i = 0; i < Operarios.size()+1; i++) {
                    cells.merge(0, valor, 1, (Vigencias.size()*3)+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS
                    int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+4), (((Vigencias.size()*i)*3)+i), ((porcionesLocal.size()+3)+16), (((Vigencias.size()*(i+1))*3)+i)+1);
                    Chart ch1 = worksheet.getCharts().get(idx1);
                    if (i == 0) { //SI EL CONTADOR ES DIFERENTE A 0 OSEA A LA PRIMERA TABLA TOTALIZADORA ENTONCES ASIGNARLE EL NOMBRE TOTAL CONSUMO 0
                        ch1.getTitle().setText("TOTAL LECTURAS LEIDAS\nTODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    } else {
                        ch1.getTitle().setText("TOTAL LECTURAS LEIDAS \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    }
                    ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA

                    columnaVIGENCIA += 1;
                    for (int j = 0; j < Vigencias.size(); j++) {
                        cells.merge(1, columnaVIGENCIA, 1, 3); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),columnaVIGENCIA).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        celdaVIGENCIAS += cellChar + "2";
                        celdaLEIDO += cellChar + (porcionesLocal.size()+4);
                        if (j < (Vigencias.size()-1)) {
                            celdaVIGENCIAS += ",";
                            celdaLEIDO += ",";
                        }
                        columnaVIGENCIA += 3;
                    }

                    String celda = "A";
                    for (int j = 0; j < Vigencias.size()*3; j++) {
                        //COLOREAR COLUMNAS PORCIONES
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        if (i != 0 && j == 0) {
                            //ASIGNAR COLOR A LAS COLUMNAS PORCION
                            cells.setColumnWidth(valor, 8.43); //CAMBIAR TAMAÑO A LA COLUMNA PORCION
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
                            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
                            flagCOLOR.setCellShading(true); //GUARDAR COLOR
                            style.setHorizontalAlignment(TextAlignmentType.LEFT); //ALINEAR A LA IZQUIERDA
                            flagCOLOR.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "3:" + cellChar + (porcionesLocal.size()+4)); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
                            celda = cellChar;
                        }

                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+3),valor).getName()).replaceAll(""+(porcionesLocal.size()+4),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+4));
                        cell.setFormula("=SUM(" + cellChar + "4:" + cellChar + (porcionesLocal.size()+3) + ")");

                    }
                    //CREAR GRAFICA 'TOTAL CONSUMO 0 X OPERARIO' Y POSICIONARLA
                    ch1.getNSeries().add(celda+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
                    ch1.getNSeries().setCategoryData(celdaVIGENCIAS); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch1.getNSeries().get(0).setName("="+celda+""+(porcionesLocal.size()+4)); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch1.getNSeries().get(0).setValues(celdaLEIDO); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

                    celdaVIGENCIAS = "=";
                    celdaLEIDO = "=";
                    valor += 1;
                }
            }

            wb.save("files\\LECTURAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            file.delete(); //ELIMINAR ARCHIVO DE .csv

            INFORME();

        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    //METODO informe -> CONSUMO_0
    public void infoCONSUMO_0() {
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = "";
            //SI LA CANTIDAD DE OPERARIOS FILTRADOS ES DIFERENTE A 0 Y A LA CANTIDAD TOTAL EXISTENTES HACER ESTO
            if (Operarios.size() != 0 && Operarios.size() != CHBX_CODOPE.length) {
                CODOPE = " AND (";
                //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
                for (int j = 0; j < Operarios.size(); j++) {
                    CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                    if (j < (Operarios.size() - 1)) {
                        CODOPE += " OR ";
                    }
                }
                CODOPE += ")";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA HACER ESTO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i)); //AGREGAR PORCIONES FILTRADAS A LA LISTA LOCAL
                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS

                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int j = 0; j < Operarios.size(); j++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(j) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int l = 0; l < Vigencias.size(); l++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_operario = '" + Operarios.get(j) + "') AND (vigencia = '" + Vigencias.get(l) + "')) AS '" + Vigencias.get(l) + ":"+ Operarios.get(j) +"'";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int j = 0; j < Vigencias.size(); j++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                        if (j+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }
            }

            //SI NO SE FILTRO NINGUNA PORCION HACER ESTO
            if (Porciones.size() == 0) {
                //CICLO QUE AGREGA TODAS LAS PORCIONES EXISTENTES EN UNA LISTA LOCAL
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }

                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS

                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int i = 0; i < Operarios.size(); i++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(i) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0) AND (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":"+ Operarios.get(i) +"'";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";

                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int i = 0; i < Vigencias.size(); i++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant = 0) AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + "'";
                        if (i+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";
                }
            }

            List<String> resultLIST = new ArrayList(); //LISTA PARA SACAR LOS RESULTADOS DE CADA FILA

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String datosXporcion = "";

                //SI NO SE FILTRO NINGUN OPERARIO O SE FILTRO MAS DE UN OPERARIO HACER ESTO
                if (Operarios.size() == 0 || Operarios.size() > 1) {
                    //EN TOTAL = CODIGO PORCION x VIGENCIAS -> RESULTADO
                    String result = rs.getString("codigo_porcion");
                    datosXporcion += result + ",";
                    for (int i = 0; i < Vigencias.size(); i++) {
                        result = rs.getString(Vigencias.get(i));
                        if (Operarios.size() == 0) {
                            datosXporcion += result;
                            if (i < (Vigencias.size()-1)) {
                                datosXporcion += ",";
                            }
                        } else {
                            datosXporcion += result + ",";
                        }
                    }
                }

                //CICLO POR OPERARIO = CODIGO_PORCION x VIGENCIAS -> RESULTADO
                for (int i = 0; i < Operarios.size(); i++) {
                    String result = rs.getString("codigo_porcion:" + Operarios.get(i));
                    datosXporcion += result + ",";
                    for (int j = 0; j < Vigencias.size(); j++) {
                        result = rs.getString(Vigencias.get(j) + ":" + Operarios.get(i));
                        datosXporcion += result;
                        if (j < Vigencias.size()-1 || i < Operarios.size()-1) {
                            datosXporcion += ",";
                        }
                    }
                }
                resultLIST.add(datosXporcion);
            }

            con.close(); //CERRAR CONEXION

            File file = new File("files\\CONSUMO_0.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter write = new PrintWriter(file); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO


            String estructura = ""; //ESTRUCTURA PRIMERA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0) {
                estructura += "TODOS LOS OPERARIOS"; //TOTAL
            } else if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                estructura += "TODOS LOS OPERARIOS FILTRADOS"; //TOTAL
                //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    estructura += ",";
                }
            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "OPERARIO " + Operarios.get(i);
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            estructura = ""; //VACIAR EL STRING

            //ESCRIBIR LAS PORCIONES Y LAS VIGENCIAS EN LA SEGUNDA FILA DE LA ESTRUCTURA
            int OyV; //ENTERO QUE SERVIRA PARA LA LONGITUD DEL CICLO
            //SI SE FILTRO SOLAMENTE 1 OPERARIO
            if (Operarios.size() == 1) {
                OyV = 1; //SOLAMENTE REPETIR EL CICLO 1 VEZ
            } else {
                OyV = Operarios.size() + 1;  //PORCIONES SELECCIONADAS + 1 DEL TOTAL
            }

            for (int i = 0; i < OyV; i++) { //CICLO POR CADA OPERARIO QUE EXISTA AGREGAR LAS VIGENCIAS EXISTENTES
                estructura += "PORCION,";
                for (int j = 0; j < Vigencias.size(); j++) {
                    estructura += ("VIG" + Vigencias.get(j));
                    if (j < (Vigencias.size()-1)) { //SI j ES MENOR AL TOTAL DE VIGENCIAS, SEPARAR LAS VIGENCIAS HASTA SER IGUAL AL TOTAL DE VIGENCIAS, ES DECIR, HASTA QUE TERMINE DE SEPARAR TODAS LAS VIGENCIAS
                        estructura += ",";
                    }
                }
                if (Operarios.size() > 1 && i < (Operarios.size())) { //SI SE FILTRO MAS DE UN OPERARIO Y j ES MENOR A CADA OPERARIO SEPARAR TODA LA ESTRUCTURA PARA VOLVER A REESCRIBIR LAS PORCIONES Y VIGENCIAS DE CADA OPERARIO HASTA QUE j SEA IGUAL, ES DECIR, TERMINE DE SEPARAR TODOS LOS OPERARIOS
                    estructura += ",";
                }
            }
            write.println(estructura);

            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                write.println(resultLIST.get(i));
            }
            //AÑADIR TOTALIZADOR
            estructura = ""; //ESTRUCTURA ULTIMA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0 || Operarios.size() > 1) {
                estructura += "TOTAL"; //TOTAL
                if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                    //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                    for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                        estructura += ",";
                    }
                }

            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "TOTAL";
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            write.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wb = new Workbook("files\\CONSUMO_0.csv"); //NUEVO LIBRO
            Worksheet worksheet = wb.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (worksheet.getCells().getCell(0,worksheet.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = worksheet.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA PORCION

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wb.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = worksheet.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 255, 0)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A1:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS SEGUNDAS FILAS Y COLUMNAS PORCION
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = worksheet.getCells().createRange("A2:A"+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN HORIZONTAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = worksheet.getCells().createRange("B2:"+lastCell+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);
            range = worksheet.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS

            Cell cell;
            int valor = 0;

            //SI NO SE FILTRO NINGUN OPERARIO O SOLO SE FILTRO 1 SOLAMENTE HACER ESTO
            if (Operarios.size() <= 1) {
                for (int i = 0; i < 1; i++) {
                    for (int j = 0; j < Vigencias.size(); j++) {
                        cells.merge(0, 0, 1, Vigencias.size()+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+3));
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + (porcionesLocal.size()+2) + ")");
                    }
                    valor += 1;
                }
                //CREAR GRAFICA 'TOTAL CONSUMO 0' Y POSICIONARLA
                int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), 0, ((porcionesLocal.size()+3)+16), (Vigencias.size()+1));
                Chart ch1 = worksheet.getCharts().get(idx1);
                ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch1.getNSeries().add("A"+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData("=B2:" + lastCell + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                if (Operarios.size() == 0) {
                    ch1.getNSeries().get(0).setName("=\"TOTAL CONSUMO 0 LECTURA\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch1.getNSeries().get(0).setName("=\"TOTAL CONSUMO 0 LECTURA\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }

                ch1.getNSeries().get(0).setValues("=B"+(porcionesLocal.size()+3)+":" + lastCell + +(porcionesLocal.size()+3)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
            } else { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                for (int i = 0; i < Operarios.size()+1; i++) {
                    cells.merge(0, valor, 1, Vigencias.size()+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS
                    int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), (Vigencias.size()*i+i), ((porcionesLocal.size()+3)+16), (Vigencias.size()+1)*(i+1));
                    Chart ch1 = worksheet.getCharts().get(idx1);
                    if (i == 0) { //SI EL CONTADOR ES DIFERENTE A 0 OSEA A LA PRIMERA TABLA TOTALIZADORA ENTONCES ASIGNARLE EL NOMBRE TOTAL CONSUMO 0
                        ch1.getTitle().setText("TOTAL CONSUMO 0 LECTURA\nTODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    } else {
                        ch1.getTitle().setText("TOTAL CONSUMO 0 LECTURA \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    }
                    ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                    String celda = "A";
                    String columnaINICIAL = "";
                    String columnaFINAL = "";
                    for (int j = 0; j < Vigencias.size(); j++) {
                        //COLOREAR COLUMNAS PORCIONES
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        if (i != 0 && j == 0) {
                            //ASIGNAR COLOR A LAS COLUMNAS PORCION
                            cells.setColumnWidth(valor, 8.43); //CAMBIAR TAMAÑO A LA COLUMNA PORCION
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
                            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
                            flagCOLOR.setCellShading(true); //GUARDAR COLOR
                            style.setHorizontalAlignment(TextAlignmentType.LEFT); //ALINEAR A LA IZQUIERDA
                            flagCOLOR.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "3:" + cellChar + (porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
                            celda = cellChar;
                        }

                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+3));
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + (porcionesLocal.size()+2) + ")");

                        if (j == 0) {
                            columnaINICIAL = cellChar;
                        }
                        if (j == Vigencias.size()-1) {
                            columnaFINAL = cellChar;
                        }
                    }
                    //CREAR GRAFICA 'TOTAL CONSUMO 0 X OPERARIO' Y POSICIONARLA
                    ch1.getNSeries().add(celda+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
                    ch1.getNSeries().setCategoryData("="+columnaINICIAL+"2:" + columnaFINAL + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch1.getNSeries().get(0).setName("="+celda+""+(porcionesLocal.size()+3)); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch1.getNSeries().get(0).setValues("="+columnaINICIAL+""+(porcionesLocal.size()+3)+":" + columnaFINAL + +(porcionesLocal.size()+3)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

                    valor += 1;
                }
            }

            wb.save("files\\CONSUMO_0.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            file.delete(); //ELIMINAR ARCHIVO DE .csv

            INFORME();

        } catch (Exception ex) {
            dialog.dispose();
            JOptionPane.showMessageDialog(null, "ERROR: PROCESO INTERRUMPIDO. POR FAVOR, CIERRE TODAS LAS PESTAÑAS RELACIONADAS AL INFORME Y VUELTA A INTENTAR NUEVAMENTE", "",JOptionPane.INFORMATION_MESSAGE);
        }

    }

    //METODO informe -> CONSUMOS_NEGATIVOS
    public void infoCONSUMOS_NEGATIVOS() {
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = "";
            //SI LA CANTIDAD DE OPERARIOS FILTRADOS ES DIFERENTE A 0 Y A LA CANTIDAD TOTAL EXISTENTES HACER ESTO
            if (Operarios.size() != 0 && Operarios.size() != CHBX_CODOPE.length) {
                CODOPE = " AND (";
                //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
                for (int j = 0; j < Operarios.size(); j++) {
                    CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                    if (j < (Operarios.size() - 1)) {
                        CODOPE += " OR ";
                    }
                }
                CODOPE += ")";
            }

            //LISTAR PORCIONES
            ArrayList<String> porcionesLocal = new ArrayList<String>(); //LISTA LOCAL QUE TENDRA LAS MISMA CANTIDAD DE PORCIONES ESTEN FILTRADAS O NO
            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI ALGUNA PORCION ESTA FILTRADA HACER ESTO
            for (int i = 0; i < Porciones.size(); i++) {
                porcionesLocal.add(Porciones.get(i)); //AGREGAR PORCIONES FILTRADAS A LA LISTA LOCAL
                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS

                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int j = 0; j < Operarios.size(); j++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(j) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int l = 0; l < Vigencias.size(); l++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (codigo_operario = '" + Operarios.get(j) + "') AND (vigencia = '" + Vigencias.get(l) + "')) AS '" + Vigencias.get(l) + ":"+ Operarios.get(j) +"'";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int j = 0; j < Vigencias.size(); j++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0) AND (codigo_porcion = '"+porcionesLocal.get(i)+"') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                        if (j+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS WHERE (codigo_porcion = '" + Porciones.get(i) + "')";
                    if (i < (Porciones.size()-1)) {
                        query += " UNION ";
                    }
                }
            }

            //SI NO SE FILTRO NINGUNA PORCION HACER ESTO
            if (Porciones.size() == 0) {
                //CICLO QUE AGREGA TODAS LAS PORCIONES EXISTENTES EN UNA LISTA LOCAL
                for (int i = 0; i < CHBX_CODPOR.length; i++) {
                    porcionesLocal.add(CHBX_CODPOR[i].getText());
                }

                //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
                if (Operarios.size() != 0) {
                    query += "SELECT"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS

                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                        query += " codigo_porcion,";
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0)" + CODOPE + " AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + "'";
                            if (j+1 < Vigencias.size()) {
                                query += ",";
                            }
                        }
                    }

                    for (int i = 0; i < Operarios.size(); i++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                        if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                            query += ",";
                        }
                        query += " codigo_porcion AS 'codigo_porcion:" + Operarios.get(i) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                        for (int j = 0; j < Vigencias.size(); j++) {
                            query += ", COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0) AND (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')) AS '" + Vigencias.get(j) + ":"+ Operarios.get(i) +"'";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";

                }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
                else {
                    query += "SELECT codigo_porcion,";
                    for (int i = 0; i < Vigencias.size(); i++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                        query += " COUNT(*) FILTER(WHERE (lectura_act - lectura_ant < 0) AND (vigencia = '" + Vigencias.get(i) + "')) AS '" + Vigencias.get(i) + "'";
                        if (i+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                    query += " FROM LECTURAS GROUP BY codigo_porcion";
                }
            }

            List<String> resultLIST = new ArrayList(); //LISTA PARA SACAR LOS RESULTADOS DE CADA FILA

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement(query);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String datosXporcion = "";

                //SI NO SE FILTRO NINGUN OPERARIO O SE FILTRO MAS DE UN OPERARIO HACER ESTO
                if (Operarios.size() == 0 || Operarios.size() > 1) {
                    //EN TOTAL = CODIGO PORCION x VIGENCIAS -> RESULTADO
                    String result = rs.getString("codigo_porcion");
                    datosXporcion += result + ",";
                    for (int i = 0; i < Vigencias.size(); i++) {
                        result = rs.getString(Vigencias.get(i));
                        if (Operarios.size() == 0) {
                            datosXporcion += result;
                            if (i < (Vigencias.size()-1)) {
                                datosXporcion += ",";
                            }
                        } else {
                            datosXporcion += result + ",";
                        }
                    }
                }

                //CICLO POR OPERARIO = CODIGO_PORCION x VIGENCIAS -> RESULTADO
                for (int i = 0; i < Operarios.size(); i++) {
                    String result = rs.getString("codigo_porcion:" + Operarios.get(i));
                    datosXporcion += result + ",";
                    for (int j = 0; j < Vigencias.size(); j++) {
                        result = rs.getString(Vigencias.get(j) + ":" + Operarios.get(i));
                        datosXporcion += result;
                        if (j < Vigencias.size()-1 || i < Operarios.size()-1) {
                            datosXporcion += ",";
                        }
                    }
                }
                resultLIST.add(datosXporcion);
            }

            con.close(); //CERRAR CONEXION

            File file = new File("files\\CONSUMOS_NEGATIVOS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter write = new PrintWriter(file); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO


            String estructura = ""; //ESTRUCTURA PRIMERA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0) {
                estructura += "TODOS LOS OPERARIOS"; //TOTAL
            } else if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                estructura += "TODOS LOS OPERARIOS FILTRADOS"; //TOTAL
                //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    estructura += ",";
                }
            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "OPERARIO " + Operarios.get(i);
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            estructura = ""; //VACIAR EL STRING

            //ESCRIBIR LAS PORCIONES Y LAS VIGENCIAS EN LA SEGUNDA FILA DE LA ESTRUCTURA
            int OyV; //ENTERO QUE SERVIRA PARA LA LONGITUD DEL CICLO
            //SI SE FILTRO SOLAMENTE 1 OPERARIO
            if (Operarios.size() == 1) {
                OyV = 1; //SOLAMENTE REPETIR EL CICLO 1 VEZ
            } else {
                OyV = Operarios.size() + 1;  //PORCIONES SELECCIONADAS + 1 DEL TOTAL
            }

            for (int i = 0; i < OyV; i++) { //CICLO POR CADA OPERARIO QUE EXISTA AGREGAR LAS VIGENCIAS EXISTENTES
                estructura += "PORCION,";
                for (int j = 0; j < Vigencias.size(); j++) {
                    estructura += ("VIG" + Vigencias.get(j));
                    if (j < (Vigencias.size()-1)) { //SI j ES MENOR AL TOTAL DE VIGENCIAS, SEPARAR LAS VIGENCIAS HASTA SER IGUAL AL TOTAL DE VIGENCIAS, ES DECIR, HASTA QUE TERMINE DE SEPARAR TODAS LAS VIGENCIAS
                        estructura += ",";
                    }
                }
                if (Operarios.size() > 1 && i < (Operarios.size())) { //SI SE FILTRO MAS DE UN OPERARIO Y j ES MENOR A CADA OPERARIO SEPARAR TODA LA ESTRUCTURA PARA VOLVER A REESCRIBIR LAS PORCIONES Y VIGENCIAS DE CADA OPERARIO HASTA QUE j SEA IGUAL, ES DECIR, TERMINE DE SEPARAR TODOS LOS OPERARIOS
                    estructura += ",";
                }
            }
            write.println(estructura);

            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < porcionesLocal.size(); i++) {
                write.println(resultLIST.get(i));
            }
            //AÑADIR TOTALIZADOR
            estructura = ""; //ESTRUCTURA ULTIMA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0 || Operarios.size() > 1) {
                estructura += "TOTAL"; //TOTAL
                if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                    //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                    for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                        estructura += ",";
                    }
                }
            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "TOTAL";
                for (int j = 0; j < Vigencias.size()+1; j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            write.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO
            Workbook wb = new Workbook("files\\CONSUMOS_NEGATIVOS.csv"); //NUEVO LIBRO
            Worksheet worksheet = wb.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (worksheet.getCells().getCell(0,worksheet.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = worksheet.getCells();
            cells.setColumnWidth(0, 8.43); //COLUMNA PORCION

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wb.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = worksheet.getCells().createRange("A1:"+lastCell+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 255, 0)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A1:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS SEGUNDAS FILAS Y COLUMNAS PORCION
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = worksheet.getCells().createRange("A2:A"+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN HORIZONTAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = worksheet.getCells().createRange("B2:"+lastCell+(porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);
            range = worksheet.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS

            Cell cell;
            int valor = 0;

            //SI NO SE FILTRO NINGUN OPERARIO O SOLO SE FILTRO 1 SOLAMENTE HACER ESTO
            if (Operarios.size() <= 1) {
                for (int i = 0; i < 1; i++) {
                    for (int j = 0; j < Vigencias.size(); j++) {
                        cells.merge(0, 0, 1, Vigencias.size()+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+3));
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + (porcionesLocal.size()+2) + ")");
                    }
                    valor += 1;
                }
                //CREAR GRAFICA 'TOTAL CONSUMO 0' Y POSICIONARLA
                int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), 0, ((porcionesLocal.size()+3)+16), (Vigencias.size()+1));
                Chart ch1 = worksheet.getCharts().get(idx1);
                ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch1.getNSeries().add("A"+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData("=B2:" + lastCell + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                if (Operarios.size() == 0) {
                    ch1.getNSeries().get(0).setName("=\"TOTAL CONSUMOS NEGATIVOS LECTURA\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch1.getNSeries().get(0).setName("=\"TOTAL CONSUMOS NEGATIVOS LECTURA\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }
                ch1.getNSeries().get(0).setValues("=B"+(porcionesLocal.size()+3)+":" + lastCell + +(porcionesLocal.size()+3)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
            } else { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                for (int i = 0; i < Operarios.size()+1; i++) {
                    cells.merge(0, valor, 1, Vigencias.size()+1); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS
                    int idx1 = worksheet.getCharts().add(ChartType.LINE, (porcionesLocal.size()+3), (Vigencias.size()*i+i), ((porcionesLocal.size()+3)+16), (Vigencias.size()+1)*(i+1));
                    Chart ch1 = worksheet.getCharts().get(idx1);
                    if (i == 0) { //SI EL CONTADOR ES DIFERENTE A 0 OSEA A LA PRIMERA TABLA TOTALIZADORA ENTONCES ASIGNARLE EL NOMBRE TOTAL CONSUMO 0
                        ch1.getTitle().setText("TOTAL CONSUMOS NEGATIVOS LECTURA\nTODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    } else {
                        ch1.getTitle().setText("TOTAL CONSUMOS NEGATIVOS LECTURA \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    }
                    ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                    String celda = "A";
                    String columnaINICIAL = "";
                    String columnaFINAL = "";
                    for (int j = 0; j < Vigencias.size(); j++) {
                        //COLOREAR COLUMNAS PORCIONES
                        String cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        if (i != 0 && j == 0) {
                            //ASIGNAR COLOR A LAS COLUMNAS PORCION
                            cells.setColumnWidth(valor, 8.43); //CAMBIAR TAMAÑO A LA COLUMNA PORCION
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
                            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
                            flagCOLOR.setCellShading(true); //GUARDAR COLOR
                            style.setHorizontalAlignment(TextAlignmentType.LEFT); //ALINEAR A LA IZQUIERDA
                            flagCOLOR.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "3:" + cellChar + (porcionesLocal.size()+3)); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
                            celda = cellChar;
                        }

                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        cellChar = (worksheet.getCells().getCell((porcionesLocal.size()+2),valor).getName()).replaceAll(""+(porcionesLocal.size()+3),"");
                        cell = worksheet.getCells().get(cellChar + (porcionesLocal.size()+3));
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + (porcionesLocal.size()+2) + ")");

                        if (j == 0) {
                            columnaINICIAL = cellChar;
                        }
                        if (j == Vigencias.size()-1) {
                            columnaFINAL = cellChar;
                        }
                    }
                    //CREAR GRAFICA 'TOTAL CONSUMOS NEGATIVOS X OPERARIO' Y POSICIONARLA
                    ch1.getNSeries().add(celda+(porcionesLocal.size()+1), true); //AGREGA LA SERIE
                    ch1.getNSeries().setCategoryData("="+columnaINICIAL+"2:" + columnaFINAL + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch1.getNSeries().get(0).setName("="+celda+""+(porcionesLocal.size()+3)); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch1.getNSeries().get(0).setValues("="+columnaINICIAL+""+(porcionesLocal.size()+3)+":" + columnaFINAL + +(porcionesLocal.size()+3)); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

                    valor += 1;
                }
            }

            wb.save("files\\CONSUMOS_NEGATIVOS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            file.delete(); //ELIMINAR ARCHIVO DE .csv

            INFORME();

        } catch (Exception ex) {
            dialog.dispose();
            JOptionPane.showMessageDialog(null, "ERROR: PROCESO INTERRUMPIDO. POR FAVOR, CIERRE TODAS LAS PESTAÑAS RELACIONADAS AL INFORME Y VUELTA A INTENTAR NUEVAMENTE", "",JOptionPane.INFORMATION_MESSAGE);
        }
    }

    //METODO informe -> ANOMALIAS
    public void infoANOMALIAS() {
        DATABASE sql = new DATABASE(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION
        try {
            //LISTAR OPERARIOS
            String CODOPE = "";
            //SI LA CANTIDAD DE OPERARIOS FILTRADOS ES DIFERENTE A 0 Y A LA CANTIDAD TOTAL EXISTENTES HACER ESTO
            if (Operarios.size() != 0) {
                CODOPE = " AND (";
                //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY
                for (int j = 0; j < Operarios.size(); j++) {
                    CODOPE += "codigo_operario = '" + Operarios.get(j) + "'";
                    if (j < (Operarios.size() - 1)) {
                        CODOPE += " OR ";
                    }
                }
                CODOPE += ")";
            }

            //LISTAR PORCIONES
            String CODPOR = "";
            String namePORCIONES = "";
            //SI LA CANTIDAD DE PORCIONES FILTRADOS ES DIFERENTE A 0 Y A LA CANTIDAD TOTAL EXISTENTES HACER ESTO
            if (Porciones.size() != 0 && Porciones.size() != CHBX_CODPOR.length) {
                CODPOR = " AND (";
                if (Porciones.size() == 1) {
                    namePORCIONES = "\nPORCION ";
                } else {
                    namePORCIONES = "\nPORCIONES ";
                }

                //SI HAY OPERARIOS FILTRADOS CREAR UNA PARTE DEL QUERY Y LISTAR LAS PORCIONES EN LA LISTA LOCAL
                for (int j = 0; j < Porciones.size(); j++) {
                    CODPOR += "codigo_porcion = '" + Porciones.get(j) + "'";
                    namePORCIONES += Porciones.get(j);
                    if (j < (Porciones.size() - 1)) {
                        CODPOR += " OR ";
                        namePORCIONES += " - ";
                    }
                }
                CODPOR += ")";
            }

            String query = ""; //CREAR EL QUERY DEPENDIENDO SI HAY O NO HAY FILTROS
            //SI SE FILTRO ALGUN OPERARIO, HACER ESTO
            if (Operarios.size() != 0) {
                if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO SACAR TOTAL DE TODOS LOS SELECCIONADOS
                    for (int j = 0; j < Vigencias.size(); j++) {
                        query += " COUNT(anomalia_1) FILTER(WHERE (vigencia = '"+Vigencias.get(j)+"')"+CODPOR+CODOPE+") AS '"+Vigencias.get(j)+"'";
                        if (j+1 < Vigencias.size()) {
                            query += ",";
                        }
                    }
                }

                for (int i = 0; i < Operarios.size(); i++) { //CICLO QUE GENERA UN QUERY CON TODOS LOS OPERARIOS SELECCIONADOS 1..*
                    if (Operarios.size() != 1) { //SI SE FILTRO MAS DE UNO IR SEPARANDO EL QUERY CON COMAS PARA SACAR TODOS LOS OPERARIOS FILTRADOS
                        query += ",";
                    }
                    query += " ANOMALIAS.ANOM AS 'ANOM:" + Operarios.get(i) + "', ANOMALIAS.DESCRIPCION AS 'DESCRIPCION:" + Operarios.get(i) + "'"; //QUERY CON TODAS LAS PORCIONES PERO CON SOLO LOS OPERARIOS FILTRADOS
                    for (int j = 0; j < Vigencias.size(); j++) {
                        query += ", COUNT(anomalia_1) FILTER(WHERE (codigo_operario = '" + Operarios.get(i) + "') AND (vigencia = '" + Vigencias.get(j) + "')"+CODPOR+") AS '" + Vigencias.get(j) + ":"+ Operarios.get(i) +"'";
                    }
                }

            }   //SI NO SE FILTRO NINGUN OPERARIO HACER ESTO
            else {
                query = "";
                for (int i = 0; i < Vigencias.size(); i++) { //CICLO QUE SACA TODOS LOS OPERARIOS RESUMIDAMENTE
                    query += " COUNT(anomalia_1) FILTER(WHERE (vigencia = '"+Vigencias.get(i)+"')"+CODPOR+") AS '"+Vigencias.get(i)+"'";
                    if (i < Vigencias.size()-1) {
                        query += ",";
                    }
                }
            }

            List<String> resultLIST = new ArrayList(); //LISTA PARA SACAR LOS RESULTADOS DE CADA FILA

            //CONSULTA -> QUERY
            PreparedStatement ps = con.prepareStatement("SELECT ANOMALIAS.ANOM, ANOMALIAS.DESCRIPCION,"+query+" FROM ANOMALIAS INNER JOIN LECTURAS ON LECTURAS.anomalia_1=ANOMALIAS.ANOM GROUP BY anomalia_1");
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String datosXporcion = "";

                //SI NO SE FILTRO NINGUN OPERARIO O SE FILTRO MAS DE UN OPERARIO HACER ESTO
                if (Operarios.size() == 0 || Operarios.size() > 1) {
                    //EN TOTAL = ANOM x VIGENCIAS -> RESULTADO
                    String result = rs.getString("ANOM");
                    datosXporcion += result + ",";
                    result = rs.getString("DESCRIPCION");
                    datosXporcion += result + ",";
                    for (int i = 0; i < Vigencias.size(); i++) {
                        result = rs.getString(Vigencias.get(i));
                        if (Operarios.size() == 0) {
                            datosXporcion += result;
                            if (i < (Vigencias.size()-1)) {
                                datosXporcion += ",";
                            }
                        } else {
                            datosXporcion += result + ",";
                        }
                    }
                }

                //CICLO POR OPERARIO = ANOM x VIGENCIAS -> RESULTADO
                for (int i = 0; i < Operarios.size(); i++) {
                    String result = rs.getString("ANOM:" + Operarios.get(i));
                    datosXporcion += result + ",";
                    result = rs.getString("DESCRIPCION:" + Operarios.get(i));
                    datosXporcion += result + ",";

                    for (int j = 0; j < Vigencias.size(); j++) {
                        result = rs.getString(Vigencias.get(j) + ":" + Operarios.get(i));
                        datosXporcion += result;
                        if (j < Vigencias.size()-1 || i < Operarios.size()-1) {
                            datosXporcion += ",";
                        }
                    }
                }
                resultLIST.add(datosXporcion);
            }

            con.close(); //CERRAR CONEXION

            File file = new File("files\\ANOMALIAS.csv"); //ARCHIVO PARA RETORNAR TODOS LOS DATOS EN UN ARCHIVO csv
            PrintWriter write = new PrintWriter(file); //PARA ESCRIBIR TODOS LOS DATOS EN EL NUEVO ARCHIVO


            String estructura = ""; //ESTRUCTURA PRIMERA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0) {
                estructura += "TODOS LOS OPERARIOS"; //TOTAL
            } else if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                estructura += "TODOS LOS OPERARIOS FILTRADOS"; //TOTAL
                //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                for (int j = 0; j < Vigencias.size()+2; j++) { // +1 POR LA COLUMNA PORCION
                    estructura += ",";
                }
            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "OPERARIO " + Operarios.get(i);
                for (int j = 0; j < Vigencias.size()+2; j++) { // +1 POR LA COLUMNA PORCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            estructura = ""; //VACIAR EL STRING

            //ESCRIBIR LAS PORCIONES Y LAS VIGENCIAS EN LA SEGUNDA FILA DE LA ESTRUCTURA
            int OyV; //ENTERO QUE SERVIRA PARA LA LONGITUD DEL CICLO
            //SI SE FILTRO SOLAMENTE 1 OPERARIO
            if (Operarios.size() == 1) {
                OyV = 1; //SOLAMENTE REPETIR EL CICLO 1 VEZ
            } else {
                OyV = Operarios.size() + 1;  //PORCIONES SELECCIONADAS + 1 DEL TOTAL
            }

            for (int i = 0; i < OyV; i++) { //CICLO POR CADA OPERARIO QUE EXISTA AGREGAR LAS VIGENCIAS EXISTENTES
                estructura += "ANOM,DESCRIPCION,";
                for (int j = 0; j < Vigencias.size(); j++) {
                    estructura += ("VIG" + Vigencias.get(j));
                    if (j < (Vigencias.size()-1)) { //SI j ES MENOR AL TOTAL DE VIGENCIAS, SEPARAR LAS VIGENCIAS HASTA SER IGUAL AL TOTAL DE VIGENCIAS, ES DECIR, HASTA QUE TERMINE DE SEPARAR TODAS LAS VIGENCIAS
                        estructura += ",";
                    }
                }
                if (Operarios.size() > 1 && i < (Operarios.size())) { //SI SE FILTRO MAS DE UN OPERARIO Y j ES MENOR A CADA OPERARIO SEPARAR TODA LA ESTRUCTURA PARA VOLVER A REESCRIBIR LAS PORCIONES Y VIGENCIAS DE CADA OPERARIO HASTA QUE j SEA IGUAL, ES DECIR, TERMINE DE SEPARAR TODOS LOS OPERARIOS
                    estructura += ",";
                }
            }
            write.println(estructura);

            //ESCRIBIR RESULTADOS DE CONSULTA DEBAJO DE LA ESTRUCTURA - INICIA SEGUNDA FILA
            for (int i = 0; i < 26; i++) {
                write.println(resultLIST.get(i));
            }
            //AÑADIR TOTALIZADOR
            estructura = ""; //ESTRUCTURA ULTIMA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0 || Operarios.size() > 1) {
                estructura += "TOTAL"; //TOTAL
                if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                    //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                    for (int j = 0; j < Vigencias.size()+2; j++) { // +2 POR LA COLUMNA ANOM Y DESCRIPCION
                        estructura += ",";
                    }
                }

            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "TOTAL";
                for (int j = 0; j < Vigencias.size()+2; j++) { // +2 POR LA COLUMNA ANOM Y DESCRIPCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            //AÑADIR TOTALIZADOR SIN 18 Y 28
            estructura = ""; //ESTRUCTURA ULTIMA FILA TOTAL (SI SELECCIONO MAS DE UN OPERARIO) Y POR OPERARIO
            if (Operarios.size() == 0 || Operarios.size() > 1) {
                estructura += "TOTAL SIN ANOM 18 Y 28"; //TOTAL
                if (Operarios.size() > 1) { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                    //AGREGAR SEPARADORES DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS DESPUES DE LA PRIMERA CELDA -> TODOS LOS OPERARIOS
                    for (int j = 0; j < Vigencias.size()+2; j++) { // +2 POR LA COLUMNA ANOM Y DESCRIPCION
                        estructura += ",";
                    }
                }

            }
            //AGREGAR CADA OPERARIO FILTRADO TAMBIEN SEPARANDO DEPENDIENDO DE LAS VIGENCIAS SELECCIONADAS
            for (int i = 0; i < Operarios.size(); i++) { //CICLO PARA CADA OPERARIO
                estructura += "TOTAL SIN ANOM 18 Y 28";
                for (int j = 0; j < Vigencias.size()+2; j++) { // +2 POR LA COLUMNA ANOM Y DESCRIPCION
                    if (i < (Operarios.size()-1)) {
                        estructura += ",";
                    }
                }
            }
            write.println(estructura);
            write.close(); //CIERRA LA ESCRITURA DE DATOS

            //CONVERTIR EN EXCEL CON DISEÑO -> falta decorar excel
            Workbook wb = new Workbook("files\\ANOMALIAS.csv"); //NUEVO LIBRO
            Worksheet worksheet = wb.getWorksheets().get(0); //NUEVA HOJA TOMANDO LA PRIMERA HOJA DEL LIBRO

            //GUARDAR LA LETRA DE LA ULTIMA COLUMNA
            String lastCell = (worksheet.getCells().getCell(0,worksheet.getCells().getMaxDataColumn()).getName()).replaceAll("1","");

            Cells cells; //CELDAS GENERAL
            Style style; //ESTILO
            StyleFlag flag = new StyleFlag(); //BANDERA
            StyleFlag flagCOLOR = new StyleFlag(); //BANDERA
            Range range; //RANGO

            //ASIGNAR CELDA CON UN TAMAÑO DEFINIDO
            cells = worksheet.getCells();
            cells.setColumnWidth(0, 5.71); //COLUMNA PORCION
            cells.setColumnWidth(1, 20); //COLUMNA PORCION

            //INICIALIZAR LA VARIABLE CON EL LIBRO
            style = wb.createStyle();
            //ASIGNAR BORDES, TIPO DE FUENTE Y TAMAÑO DE FUENTE A LAS CELDAS
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            flag.setBorders(true); //GUARDAR BORDEO
            style.getFont().setName("Calibri"); //CAMBIAR FUENTE A CALIBRI
            flag.setFont(true); //GUARDAR TIPO DE FUENTE
            style.getFont().setSize(11); //CAMBIAR TAMAÑO DE FUENTE
            flag.setFontSize(true); //GUARDAR TAMAÑO
            range = worksheet.getCells().createRange("A1:"+lastCell+"30"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS PRIMERAS FILAS Y COLUMNAS
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 255, 0)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A1:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR COLOR A LAS SEGUNDAS FILAS Y COLUMNAS PORCION
            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
            flagCOLOR.setCellShading(true); //GUARDAR COLOR
            range = worksheet.getCells().createRange("A2:"+lastCell+"2"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            range = worksheet.getCells().createRange("A2:A30"); //RANGO DONDE SE APLICARA EL COLOR
            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
            //ASIGNAR ALINEACIONES A LAS COLUMNAS VIGENCIAS
            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR EN EL MEDIO EN HORIZONTAL
            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
            range = worksheet.getCells().createRange("C2:"+lastCell+"28"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range.setColumnWidth(10);
            range = worksheet.getCells().createRange("A1:"+lastCell+"1"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS
            range = worksheet.getCells().createRange("A29:"+lastCell+"30"); //RANGO DONDE SE APLICARA EL DISEÑO
            range.applyStyle(style, flag); //APLICAR DISEÑO AL RANGO DE CELDAS

            Cell cell;
            int valor = 1;

            //SI NO SE FILTRO NINGUN OPERARIO O SOLO SE FILTRO 1 SOLAMENTE HACER ESTO
            if (Operarios.size() <= 1) {
                for (int i = 0; i < 1; i++) {
                    for (int j = 0; j < Vigencias.size(); j++) {
                        cells.merge(0, 0, 1, Vigencias.size()+2); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS
                        cells.merge(28, 0, 1, 2); //COMBINAR Y CENTRAR TOTAL
                        cells.merge(29, 0, 1, 2); //COMBINAR Y CENTRAR TOTAL SIN ANOMALIAS 18 Y 28
                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        String cellChar = (worksheet.getCells().getCell(28,valor).getName()).replaceAll("29","");
                        cell = worksheet.getCells().get(cellChar + "29");
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + "28)");
                        cellChar = (worksheet.getCells().getCell(29,valor).getName()).replaceAll("30","");
                        cell = worksheet.getCells().get(cellChar + "30");
                        cell.setFormula("=" + cellChar + "29 - (" + cellChar + "17+"+cellChar+"26)");
                    }
                    valor += 1;
                }
                //CREAR GRAFICA 'TOTAL SIN ANOM 18 Y 28' Y POSICIONARLA
                int idx1 = worksheet.getCharts().add(ChartType.LINE, 30, 0, 42, (Vigencias.size()+2));
                Chart ch1 = worksheet.getCharts().get(idx1);
                ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch1.getNSeries().add("A29", true); //AGREGA LA SERIE
                ch1.getNSeries().setCategoryData("=C2:" + lastCell + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                if (Operarios.size() == 0) {
                    ch1.getNSeries().get(0).setName("=\"TOTAL ANOMALIAS (SIN ANOMALIA 18 Y 28) LECTURA\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch1.getNSeries().get(0).setName("=\"TOTAL ANOMALIAS (SIN ANOMALIA 18 Y 28) LECTURA\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }
                ch1.getNSeries().get(0).setValues("=C30:" + lastCell + "30"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
                //CREAR GRAFICA 'TOTAL ANOMALIAS 18' Y POSICIONARLA
                int idx2 = worksheet.getCharts().add(ChartType.LINE, 42, 0, 54, (Vigencias.size()+2));
                Chart ch2 = worksheet.getCharts().get(idx2);
                ch2.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch2.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch2.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch2.getNSeries().add("A29", true); //AGREGA LA SERIE
                ch2.getNSeries().setCategoryData("=C2:" + lastCell + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                if (Operarios.size() == 0) {
                    ch2.getNSeries().get(0).setName("=\"TOTAL ANOMALIA 18 PREDIO DESOCUPADO LECTURA\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch2.getNSeries().get(0).setName("=\"TOTAL ANOMALIA 18 PREDIO DESOCUPADO LECTURA\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }
                ch2.getNSeries().get(0).setValues("=C17:" + lastCell + "17"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                ch2.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch2.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch2.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
                //CREAR GRAFICA 'TOTAL ANOMALIAS 18' Y POSICIONARLA
                int idx3 = worksheet.getCharts().add(ChartType.LINE, 54, 0, 66, (Vigencias.size()+2));
                Chart ch3 = worksheet.getCharts().get(idx3);
                ch3.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                ch3.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                ch3.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                ch3.getNSeries().add("A29", true); //AGREGA LA SERIE
                ch3.getNSeries().setCategoryData("=C2:" + lastCell + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                if (Operarios.size() == 0) {
                    ch3.getNSeries().get(0).setName("=\"TOTAL ANOMALIA 28 PREDIO OCUPADO LECTURA\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                } else {
                    ch3.getNSeries().get(0).setName("=\"TOTAL ANOMALIA 28 PREDIO OCUPADO LECTURA\nOPERARIO " + Operarios.get(0) + "\""); //ASIGNAR NOMBRE DE LA SERIA COMO LA CELDA
                }
                ch3.getNSeries().get(0).setValues("=C26:" + lastCell + "26"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                ch3.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                ch3.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                ch3.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

            } else { //SI SE FILTRO MAS DE UN OPERARIO HACER ESTO
                for (int i = 0; i < Operarios.size()+1; i++) {
                    cells.setColumnWidth((Vigencias.size()*i+i)+i, 5.71); //COLUMNA ANOMxOPERARIO
                    cells.merge(0, (Vigencias.size()*i+i)+i, 1, Vigencias.size()+2); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS
                    cells.merge(28, (Vigencias.size()*i+i)+i, 1, 2); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS
                    cells.merge(29, (Vigencias.size()*i+i)+i, 1, 2); //COMBINAR Y CENTRAR POR LA CANTIDAD TOTAL DE VIGENCIAS Y OPERARIOS

                    int idx1 = worksheet.getCharts().add(ChartType.LINE, 30, (Vigencias.size()*i+i)+i, 42, (Vigencias.size()+2)*(i+1));
                    Chart ch1 = worksheet.getCharts().get(idx1);
                    int idx2 = worksheet.getCharts().add(ChartType.LINE, 42, (Vigencias.size()*i+i)+i, 54, (Vigencias.size()+2)*(i+1));
                    Chart ch2 = worksheet.getCharts().get(idx2);
                    int idx3 = worksheet.getCharts().add(ChartType.LINE, 54, (Vigencias.size()*i+i)+i, 66, (Vigencias.size()+2)*(i+1));
                    Chart ch3 = worksheet.getCharts().get(idx3);
                    if (i == 0) { //SI EL CONTADOR ES DIFERENTE A 0 OSEA A LA PRIMERA TABLA TOTALIZADORA ENTONCES ASIGNARLE EL NOMBRE TOTAL CONSUMO 0
                        ch1.getTitle().setText("TOTAL ANOMALIAS (SIN ANOMALIA 18 Y 28) LECTURA\n TODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                        ch2.getTitle().setText("TOTAL ANOMALIAS 18 PREDIO DESOCUPADO LECTURA\n TODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                        ch3.getTitle().setText("TOTAL ANOMALIAS 28 PREDIO OCUPADO LECTURA\n TODOS LOS OPERARIOS FILTRADOS"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    } else {
                        ch1.getTitle().setText("TOTAL ANOMALIAS (SIN ANOMALIA 18 Y 28) LECTURA \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                        ch2.getTitle().setText("TOTAL ANOMALIAS 18 PREDIO DESOCUPADO LECTURA \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                        ch3.getTitle().setText("TOTAL ANOMALIAS 28 PREDIO OCUPADO LECTURA \nOPERARIO (" + Operarios.get(i-1) +")"); //ASIGNARLE UN NOMBRE A LA GRAFICA
                    }
                    ch1.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch1.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch1.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                    ch2.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch2.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch2.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                    ch3.getTitle().getFont().setSize(15); //ASIGNARLE UN TAMAÑO LETRA
                    ch3.getTitle().getFont().setBold(true); //ASIGNARLE NEGRILLA A LA LETRA
                    ch3.setShowLegend(false); //QUITAR LEYENDA DE LA GRAFICA
                    String celda = "A";
                    String columnaINICIAL = "";
                    String columnaFINAL = "";

                    for (int j = 0; j < Vigencias.size(); j++) {
                        //COLOREAR COLUMNAS PORCIONES
                        String cellChar = (worksheet.getCells().getCell(29,valor-1).getName()).replaceAll(""+30,"");
                        if (i != 0 && j == 0) {
                            celda = cellChar;
                            cells.setColumnWidth(valor, 5.71); //CAMBIAR TAMAÑO A LA COLUMNA ANOM
                            style.setForegroundColor(com.aspose.cells.Color.fromArgb(142, 169, 219)); //CAMBIAR COLOR
                            style.setPattern(BackgroundType.SOLID); //DEFINIRLO COMO SOLIDO
                            flagCOLOR.setCellShading(true); //GUARDAR COLOR
                            style.setHorizontalAlignment(TextAlignmentType.RIGHT); //ALINEAR A LA IZQUIERDA
                            flagCOLOR.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "2:" + cellChar + "28"); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
                            style.setHorizontalAlignment(TextAlignmentType.CENTER); //ALINEAR A LA IZQUIERDA
                            flagCOLOR.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "29:" + cellChar + "30"); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flagCOLOR); //APLICAR COLOR AL RANGO DE CELDAS
                            cellChar = (worksheet.getCells().getCell(29,valor).getName()).replaceAll(""+30,"");
                            cells.setColumnWidth((Vigencias.size()*i+i)+i+1, 20); //COLUMNA DESCRIPCIONxOPERARIO
                            style.setHorizontalAlignment(TextAlignmentType.LEFT); //ALINEAR A LA IZQUIERDA
                            flag.setAlignments(true); //GUARDAR ALINEAMIENTOS
                            range = worksheet.getCells().createRange(cellChar + "2:" + cellChar + "28"); //RANGO DONDE SE APLICARA EL COLOR
                            range.applyStyle(style, flag); //APLICAR COLOR AL RANGO DE CELDAS
                        }
                        valor += 1; //SUMA PARA SACAR LA CELDA DONDE ES EL TOTAL
                        cellChar = (worksheet.getCells().getCell(29,valor).getName()).replaceAll(""+30,"");
                        cell = worksheet.getCells().get(cellChar + "29");
                        cell.setFormula("=SUM(" + cellChar + "3:" + cellChar + "28)");
                        cell = worksheet.getCells().get(cellChar + "30");
                        cell.setFormula("=" + cellChar + "29-(" + cellChar + "17+"+cellChar+"26)");

                        if (j == 0) {
                            columnaINICIAL = cellChar;
                        }
                        if (j == Vigencias.size()-1) {
                            columnaFINAL = cellChar;
                        }
                    }

                    //CREAR GRAFICA 'TOTAL CONSUMOS NEGATIVOS X OPERARIO' Y POSICIONARLA
                    ch1.getNSeries().add(celda+"30", true); //AGREGA LA SERIE
                    ch1.getNSeries().setCategoryData("="+columnaINICIAL+"2:" + columnaFINAL + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch1.getNSeries().get(0).setName("="+celda+"30"); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch1.getNSeries().get(0).setValues("="+columnaINICIAL+"30:" + columnaFINAL + "30"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch1.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch1.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch1.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
                    ch2.getNSeries().add(celda+"17", true); //AGREGA LA SERIE
                    ch2.getNSeries().setCategoryData("="+columnaINICIAL+"2:" + columnaFINAL + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch2.getNSeries().get(0).setName("=\"TOTAL ANOMALIAS 18\""); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch2.getNSeries().get(0).setValues("="+columnaINICIAL+"17:" + columnaFINAL + "17"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch2.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch2.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch2.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO
                    ch3.getNSeries().add(celda+"26", true); //AGREGA LA SERIE
                    ch3.getNSeries().setCategoryData("="+columnaINICIAL+"2:" + columnaFINAL + "2"); //SELECCIONAR COMO CATEGORIAS LAS VIGENCIAS
                    ch3.getNSeries().get(0).setName("=\"TOTAL ANOMALIAS 28\""); //ASIGNAR NOMBRE DE LA SERIE COMO LA CELDA
                    ch3.getNSeries().get(0).setValues("="+columnaINICIAL+"26:" + columnaFINAL + "26"); //SELECCIONAR LOS DATOS DE LA SERIE QUE EN ESTE CASO SERIA EL VALOR TOTAL POR CADA VIGENCIA
                    ch3.getNSeries().get(0).getDataLabels().setShowValue(true); //MOSTRAR LAS ETIQUETAS DE DATOS EN LA GRAFICA
                    ch3.getNSeries().get(0).getDataLabels().setPosition(LabelPositionType.ABOVE); //MOSTRAR LAS ETIQUETAS DE DATOS ENCIMA DE LA LINEA DE GRAFICO
                    ch3.getNSeries().get(0).getMarker().setMarkerStyle(FillType.AUTOMATIC); //MOSTRAR LOS MARCADORES EN LA LINEA DE GRAFICO

                    valor += 2;

                }
            }

            wb.save("files\\ANOMALIAS.xlsx", SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
            file.delete(); //ELIMINAR ARCHIVO DE .csv
            INFORME();

        } catch (Exception ex) {
            dialog.dispose();
            JOptionPane.showMessageDialog(null, "ERROR: PROCESO INTERRUMPIDO. POR FAVOR, CIERRE TODAS LAS PESTAÑAS RELACIONADAS AL INFORME Y VUELTA A INTENTAR NUEVAMENTE", "",JOptionPane.INFORMATION_MESSAGE);
        }
    }

    //METODO informe -> ANOMALIASxPORCION
    public void infoANOMALIASxPORCION() {
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

            INFORME();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //METODO informe -> ANOMALIASxOPERARIO
    public void infoANOMALIASxOPERARIO() {
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

            INFORME();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //METODO GENERAR INFORME
    public void INFORME() {
        valFINISH++;
        if (valFINISH == 4) {
            try {
                //CREAR EXCEL DE INFORME
                Workbook wbINFORME = new Workbook(); //NUEVO LIBRO
                //SELECCIONAR LOS ARCHIVOS CON LOS DATOS Y UNIFICARLOS EN UN SOLO ARCHIVO
                File fileEXCEL_LECTURAS = new File("files\\LECTURAS.xlsx");
                File fileEXCEL_CONSUMO_0 = new File("files\\CONSUMO_0.xlsx");
                File fileEXCEL_CONSUMOS_NEGATIVOS = new File("files\\CONSUMOS_NEGATIVOS.xlsx");
                File fileEXCEL_ANOMALIAS = new File("files\\ANOMALIAS.xlsx");
                //File fileEXCEL_ANOMALIASxPORCION = new File("files\\ANOMALIASxPORCION.xlsx");
                //File fileEXCEL_ANOMALIASxOPERARIO = new File("files\\ANOMALIASxOPERARIO.xlsx");
                Workbook wbLECTURAS = new Workbook(fileEXCEL_LECTURAS.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                Workbook wbCONSUMO_0 = new Workbook(fileEXCEL_CONSUMO_0.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                Workbook wbCONSUMOS_NEGATIVOS = new Workbook(fileEXCEL_CONSUMOS_NEGATIVOS.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                Workbook wbANOMALIAS = new Workbook(fileEXCEL_ANOMALIAS.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                //Workbook ANOMALIASxPORCION = new Workbook(fileEXCEL_ANOMALIASxPORCION.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                //Workbook ANOMALIASxOPERARIO = new Workbook(fileEXCEL_ANOMALIASxOPERARIO.getAbsolutePath()); //NUEVO LIBRO DEL ARCHIVO DE ANOMALIAS
                //COMBINAR HOJAS EN EL INFORME
                wbINFORME.combine(wbLECTURAS);
                wbINFORME.combine(wbCONSUMO_0);
                wbINFORME.combine(wbCONSUMOS_NEGATIVOS);
                wbINFORME.combine(wbANOMALIAS);
                //wbINFORME.combine(ANOMALIASxPORCION);
                //wbINFORME.combine(ANOMALIASxOPERARIO);
                wbINFORME.getWorksheets().removeAt(0); //ELIMINAR LA PRIMERA HOJA VACIA DEL LIBRO
                wbINFORME.save("files\\INFORME.xlsx");
                //ELIMINAR LIBROS COPIADOS
                fileEXCEL_LECTURAS.delete();
                fileEXCEL_CONSUMO_0.delete();
                fileEXCEL_CONSUMOS_NEGATIVOS.delete();
                fileEXCEL_ANOMALIAS.delete();
                //fileEXCEL_ANOMALIASxPORCION.delete();
                //fileEXCEL_ANOMALIASxOPERARIO.delete();

                dialog.dispose(); //CERRAR LOADING
                JOptionPane.showMessageDialog(null, "SE EXPORTO CORRECTAMENTE EL INFORME", "",JOptionPane.INFORMATION_MESSAGE);
                File ARCHIVOS = new File("files");
                Runtime.getRuntime().exec("cmd /c start " + ARCHIVOS.getAbsolutePath() + " && exit");
                valFINISH = 0;
            } catch (Exception ex) {
                ex.printStackTrace();
                dialog.dispose();
                JOptionPane.showMessageDialog(null, "ERROR: PROCESO INTERRUMPIDO. POR FAVOR, CIERRE TODAS LAS PESTAÑAS RELACIONADAS AL INFORME Y VUELTA A INTENTAR NUEVAMENTE", "",JOptionPane.INFORMATION_MESSAGE);
            }
        }
    }

    //METODO MAIN
    public static void main(String[] args) {
        new PROGRAMA();
    }

}

/* @author marxs7
 * basado en código de Chuidiang
 * En este ejemplo se usa Gmail pero pueden utilizar cualquier servidor
 * solo remplacen los valores asignados a servidorSMTP Y puertoEnvio.
 */
import java.util.List;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;



////////////importacion de excel







import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

///



public class EnviarCorreoMejorado
{   String miCorreo;
    String miPassword;
    String servidorSMTP="smtp.gmail.com";
    String puertoEnvio="587";// puertoEnvio="587";//465 para mensajes normales y 587 con adjuntos
    String[] destinatarios;
    String asunto;
    String cuerpo = null;//cuerpo del mensaje
    String[] archivoAdjunto;
       
    static String [][] basededatosPa;
    static int numeroTotalDeFilas;
    static int numeroTotalDeColumnas;
    static String CadenaS = "";
    
  public EnviarCorreoMejorado(String usuario,String pass,String[] dest,String asun,String mens,String[] archivo){
        destinatarios=dest;
        asunto=asun;
        cuerpo=mens;
        archivoAdjunto=archivo;  
        miCorreo=usuario;
        miPassword=pass;
     }
  public static void muestraContenido(String archivo) throws FileNotFoundException, IOException {
      String cadena;
      FileReader f = new FileReader(archivo);
      BufferedReader b = new BufferedReader(f);
      while((cadena = b.readLine())!=null) 
      {
      	
      	CadenaS = CadenaS + cadena + "\n";
      }
      
      b.close();
  }  
  
    public void Enviar() throws MessagingException
    {
            Properties props=null;
            props = new Properties();
            props.put("mail.smtp.host", servidorSMTP);
            props.setProperty("mail.smtp.starttls.enable", "true");//si usa Yahoo comentar esta linea         
            props.setProperty("mail.smtp.port", "587");
            props.setProperty("mail.smtp.user", miCorreo);
            props.setProperty("mail.smtp.auth", "true");
            SecurityManager security = System.getSecurityManager();
     
            Authenticator auth = new EnviarCorreoMejorado.autentificadorSMTP();
            Session session = Session.getInstance(props, auth);
           // session.setDebug(true); //Descomentar para ver el proceso de envio detalladamente
            
            // Se compone la parte del texto
            BodyPart texto = new MimeBodyPart();
            texto.setText(cuerpo);
            
            // Se compone el adjunto 
            BodyPart[] adjunto=new BodyPart[archivoAdjunto.length];
            for(int j=0;j<archivoAdjunto.length;j++){
            adjunto[j]=new MimeBodyPart();
            adjunto[j].setDataHandler(
                new DataHandler(new FileDataSource(archivoAdjunto[j])));
            
            String[] rutaArchivo = archivoAdjunto[j].split("/");//separamos las palabras que forman la url y las                 ponemos en arreglo  de cadenas
            int nombre=rutaArchivo.length-1;//del array buscamos la ultima posicion
            adjunto[j].setFileName(rutaArchivo[nombre]);//la ultima posicion debe tener el nombre del archivo
            }
            
            // Una MultiParte para agrupar texto e imagen.
            MimeMultipart multiParte = new MimeMultipart();
            multiParte.addBodyPart(texto);
            for(BodyPart aux:adjunto){
                multiParte.addBodyPart(aux);
            }
            
            // Se compone el correo, dando to, from, subject y el
            // contenido.
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(miCorreo));
            Address []destinos = new Address[destinatarios.length];//Aqui usamos el arreglo de destinatarios
            for(int i=0;i<destinos.length;i++){
                destinos[i]=new InternetAddress(destinatarios[i]);
            }
            message.addRecipients(Message.RecipientType.TO, destinos);//agregamos los destinatarios
            message.setSubject(asunto);
            message.setContent(multiParte);

            // Se envia el correo.
            Transport t = session.getTransport("smtp");
            t.connect(miCorreo, miPassword);
            t.sendMessage(message, message.getAllRecipients());
            System.out.println("Correo Enviado exitosamente!"); 
            t.close();    
        }
     private class autentificadorSMTP extends javax.mail.Authenticator {
        public PasswordAuthentication getPasswordAuthentication() {
            return new PasswordAuthentication(miCorreo, miPassword);
        }
    }
     public static void main(String[] args) throws MessagingException, Exception {
    	 /**
    	  * 
    	  /cadena en array 
    	 String colores = "rojo,amarillo,verde,azul,morado,marrón";
    	 String[] arrayColores = colores.split(",");
    	  
    	 // En este momento tenemos un array en el que cada elemento es un color.
    	 for (int i = 0; i < arrayColores.length; i++) {
    	 	System.out.println(arrayColores[i]);
    	 }
    	 */
    	 ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    	 ///excel
    	 
    	 //
 		// An excel file name. You can create a file name with a full
 		// path information.
 		//
    	 
    	//ruta de archivo  
    	 
 		String filename = "bd/base.xls";
 		//
 		// Create an ArrayList to store the data read from excel sheet.
 		//
 		List sheetData = new ArrayList();
 		FileInputStream fis = null;
 		try {
 			//
 			// Create a FileInputStream that will be use to read the
 			// excel file.
 			//
 			fis = new FileInputStream(filename);
 			//
 			// Create an excel workbook from the file system.
 			//
 			HSSFWorkbook workbook = new HSSFWorkbook(fis);
 			//
 			// Get the first sheet on the workbook.
 			//
 			HSSFSheet sheet = workbook.getSheetAt(0);
 			//
 			// When we have a sheet object in hand we can iterator on
 			// each sheet's rows and on each row's cells. We store the
 			// data read on an ArrayList so that we can printed the
 			// content of the excel to the console.
 			//
 			Iterator rows = sheet.rowIterator();
 			while (rows.hasNext()) {
 				HSSFRow row = (HSSFRow) rows.next();
 				
 				Iterator cells = row.cellIterator();
 				List data = new ArrayList();
 				while (cells.hasNext()) {
 					HSSFCell cell = (HSSFCell) cells.next();
 				//	System.out.println("Añadiendo Celda: " + cell.toString());
 					data.add(cell);
 				}
 				sheetData.add(data);
 			}
 		} catch (IOException e) {
 			e.printStackTrace();
 		} finally {
 			if (fis != null) {
 				fis.close();
 			}
 		}
 		showExelData(sheetData);
 		
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
 		
 	 String Correo="dee.quintanaroo2@gmail.com";	
 	 
 	 String Contrasenia="seee23qroo";
   	 
 	 String Asunto = "ENVÍA RESULTADOS PLANEA 2015";
   	 
   	 String [] archivosParaAjuntar = new String[5];
 
   	 muestraContenido("txt/archivo.txt");
  	
   	 String T  = CadenaS;
   	 //no cambia
   	 
   	 String JS = "JS";
   	 
   	 String []  correoAenviar = new String[6];
   	 
   	 //apartir del 62 debe de ocurrir unerrror porque no tengo el correo 
   	 
   	 //numero de correo es para llevar el control de donde nos quedamos
   	 //numeroTotalDeFilas es la variable que quite

   	 
   	 
   	 for(int numeroDeCorreo= 88; numeroDeCorreo<99; numeroDeCorreo++ )
   	 {	
   		//numeroDeCorreo
   		if(JS.equals(basededatosPa[numeroDeCorreo][2]))
   		{
   			for(int xs= 0; xs<3; xs++ )
   			{	
   				correoAenviar[xs] = basededatosPa[numeroDeCorreo][5+xs];
   				
        	}
   		}else{
   		
        for(int xs= 3; xs<6; xs++ )
      	 {	
        	correoAenviar[xs] = basededatosPa[numeroDeCorreo][2+xs];
        
      	 }
   		
   	  	 archivosParaAjuntar[0]="archivos/Lenguaje y comunicación Secundaria_2015 Niveles de logro.pdf";
   	  	 archivosParaAjuntar[1]="archivos/Matemáticas Secundaria_2015 Niveles de logro.pdf";
   	  	 
   	  	 double d = Double.parseDouble(basededatosPa[numeroDeCorreo][2]);
   	  	 int i = (int) d;
   	  	 
   	  	 String enteroString = Integer.toString(i);
   	  	 if(i<10)
   	  	 {
   	  		 enteroString = "0" + enteroString;
   	  	 }
   	  	 
   	  
   	  	 
   	  	 

   	  	 

   	  	 
   	  	 
   	  	 
   	  	 
   	  	 archivosParaAjuntar[2]="archivos/PLANEA_BASICA_2015_TELESECUNDARIA/"
   	  			 +"TELESECUNDARIA ZE"+enteroString+"/"
   	  			 +"planea_secu_alumnos_ze_0"+enteroString+".xlsx";
   	  	 
   	  	 
   	  	 
   	 
   	  	    	  	 
   	  	 archivosParaAjuntar[3]="archivos/PLANEA_BASICA_2015_TELESECUNDARIA/"
   	  			 +"TELESECUNDARIA ZE"+enteroString+"/"
   	  			 +"planea_secu_ze_0"+enteroString+".xlsx";
   
   	  	 
   	  	 archivosParaAjuntar[4]="archivos/PLANEA_BASICA_2015_TELESECUNDARIA/"
   	  			 +"TELESECUNDARIA ZE"+enteroString+"/"
   	  			  +"PLANEA EB RESULTADOS 2015 Secundaria ZE"+enteroString+"-Teles.pptx";   	 
  
     	/*
   		 
 		  C. Supervisor de la Zona Escolar 001
			Primaria General
		*/
  	  	 
  	  	 String Texto = "C. Supervisor de la Zona Escolar "+enteroString+"\n\n"+basededatosPa[numeroDeCorreo][1]+"\n"+T;
   	   	 
  	  	 
  	  	 EnviarCorreoMejorado ai = new EnviarCorreoMejorado(Correo,Contrasenia, correoAenviar, Asunto, Texto, archivosParaAjuntar);  
   	   	 ai.Enviar();
   	  	 
   	   	 System.out.println("Terminado "+numeroDeCorreo);  	
   		};
   	 };
    	 System.out.println("Terminado el envio de archivos");
    	
     }
     
     
     ///parte de la lectura 
     
 	private static void showExelData(List sheetData) {
		//
		// Iterates the data and print it out to the console.
		//	
		//System.out.println(sheetData.size());
		List columna = (List) sheetData.get(1);
		//System.out.println(list.size());
		
		basededatosPa = new String[sheetData.size()][columna.size()];
		
		numeroTotalDeColumnas = columna.size();
		numeroTotalDeFilas = sheetData.size();
		

		
		for (int i = 0; i < numeroTotalDeFilas; i++) {
			List list = (List) sheetData.get(i);
			for (int j = 0; j < list.size(); j++) 
			{
				Cell cell = (Cell) list.get(j);
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				{
					
					basededatosPa [i][j] = Double.toString(cell.getNumericCellValue());
				} 
				else if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
				{
					basededatosPa [i][j] = cell.getStringCellValue();							
				} 
			}
		}
		
		
	}
 	
 	
 	////////////////////////////////////////////////////////////////////////////////////////////////////////
    
        
    }

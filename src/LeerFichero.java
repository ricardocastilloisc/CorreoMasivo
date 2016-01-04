import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
 
public class LeerFichero {
	
	static String CadenaS = "";
	
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
 
    public static void main(String[] args) throws IOException {
        muestraContenido("txt/archivo.txt");
        System.out.println(CadenaS);
        
    }
   
}
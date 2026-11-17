import java.io.*;
//import java.io.Writer;

public class Testing {
    public static void main(String[] args){

      try
      {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(new File("data.txt")), java.nio.charset.StandardCharsets.UTF_8));         
        
        
        char[] temp = new char[10000];
        int len;
        StringBuffer b  = new StringBuffer();
        while((len=reader.read(temp))!=-1)
        { //end of stream is not reached
          b.append(temp,0,len);
        }
        
        //String s = reader.readLine();
        String s = b.toString();
        PrintStream out = new PrintStream(System.out, true, "UTF-16");
        
        out.println(s);
        
        reader.close();
        
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File("test.txt")), java.nio.charset.StandardCharsets.UTF_8));
        //BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File("test.txt"))));
        s = s+"\u4e2d";
        
        byte[] bytes = s.getBytes(java.nio.charset.StandardCharsets.UTF_8);
        String utf8EncodedString = new String(bytes,java.nio.charset.StandardCharsets.UTF_8);
        System.out.println(utf8EncodedString);
        
        //writer.write(utf8EncodedString);writer.newLine();
        writer.write(s);writer.newLine();
        writer.flush();
        writer.close();
        
        
        } 
       catch (Exception e) 
       {
        System.out.println(e.getMessage());
       } 
    }    
}
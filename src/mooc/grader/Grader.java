/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package mooc.grader;

/**
 *
 * @author aavendan
 */
public class Grader {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        try {
            //String original = "docx/Avendaño Allan-seccionWORD_V1R.docx";
            //String original = "docx/Avendaño Allan-seccionWORD_V2.docx";
            String original = "docx/original.docx";
            String respuesta = "docx/response-2.docx";
            Verifier v = new Verifier();
            v.setFileName(Verifier.INDEX_ORIGINAL, original);
            v.setFileName(Verifier.INDEX_RESPONSE, respuesta);
            v.validate();
            System.gc();
        } catch (Throwable t) {
            t.printStackTrace();
        }// TODO code application logic here
    }
    
    public static void p(Object o) {
        System.out.println(o);
    }

    
    
}

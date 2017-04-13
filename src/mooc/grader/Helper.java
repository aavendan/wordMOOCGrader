/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

import java.util.Iterator;
import java.util.List;
import javax.xml.bind.JAXBElement;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 *
 * @author aavendan
 */
public class Helper {

    public static String getTextFromR(List<Object> lp) {
        String str = "";

        for (Object o2 : lp) {
            JAXBElement element = (JAXBElement) o2;
            if (element.getValue() instanceof org.docx4j.wml.Text) {
                Text text = (org.docx4j.wml.Text) element.getValue();
                str += text.getValue();
            }
        }

        return str;
    }

    public static String getTextFromP(List<Object> lp) {
        String str = "";

        for (Object olp : lp) {
            if (olp instanceof org.docx4j.wml.R) {
                R r = (org.docx4j.wml.R) olp;
                List<Object> c2 = r.getContent();

                for (Object o2 : c2) {
                    if (o2 instanceof javax.xml.bind.JAXBElement) {
                        JAXBElement element = (JAXBElement) o2;
                        if (element.getValue() instanceof org.docx4j.wml.Text) {
                            Text text = (org.docx4j.wml.Text) element.getValue();
                            str += text.getValue();
                        }
                    }
                }
            }
        }

        return str;
    }

    public static String getTextFromFtr(org.docx4j.wml.Ftr ftr) {
        String text = null;

        List<Object> content = ftr.getContent();

        if (content.size() > 0) {
            text = new String();
            Iterator<Object> it = content.iterator();
            while (it.hasNext()) {
                Object obj = it.next();
                if (obj instanceof org.docx4j.wml.P) {
                    P p = (org.docx4j.wml.P) obj;
                    if (p.getPPr().getFramePr() == null) {
                        text += getTextFromP(p.getContent());
                    }
                }
            }
        }

        return text;
    }

    public static boolean isNumbered(org.docx4j.wml.Ftr ftr) {
        return tillPage(ftr) != -1;
    }

    public static int tillPage(org.docx4j.wml.Ftr ftr) {
        String text = null;

        List<Object> content = ftr.getContent();

        if (content.size() > 0) {
            text = new String();
            Iterator<Object> it = content.iterator();
            while (it.hasNext()) {
                Object obj = it.next();
                if (obj instanceof org.docx4j.wml.P) {
                    P p = (org.docx4j.wml.P) obj;
                    if (p.getPPr().getFramePr() != null) {
                        text += getTextFromP(p.getContent());
                    }
                }
            }

            text = text.trim().replace("PAGE", "").trim();
        }

        int page;

        try {
            page = Integer.valueOf(text);
        } catch (Exception ex) {
            return -1;
        }

        return page;
    }

    public static boolean compareTo(Object s1, Object s2) {
        if (s1 == null || s2 == null) {
            return false;
        }

        //System.out.println("  *** "+String.valueOf(s1)+" "+String.valueOf(s2).toString()+" "+(String.valueOf(s1).compareTo(String.valueOf(s2).toString()) == 0));
        return String.valueOf(s1).compareTo(String.valueOf(s2).toString()) == 0;
    }

    public static boolean isHeading(String type, String s1, String s2) {
        return s1.compareTo(type) == 0 || s2.compareTo(type) == 0;
    }

    public static String escapeNonAscii(String str) {

        StringBuilder retStr = new StringBuilder();
        for (int i = 0; i < str.length(); i++) {
            int cp = Character.codePointAt(str, i);
            int charCount = Character.charCount(cp);
            if (charCount > 1) {
                i += charCount - 1; // 2.
                if (i >= str.length()) {
                    throw new IllegalArgumentException("truncated unexpectedly");
                }
            }

            if (cp < 128) {
                retStr.appendCodePoint(cp);
            } else {
                retStr.append(String.format("\\u%x", cp));
            }
        }
        return retStr.toString();
    }

    public static boolean similarTo(String nC1, String nC2, double X) {
        int color1 = (int) Long.parseLong(nC1, 16);
        int r1 = (color1 >> 16) & 0xFF;
        int g1 = (color1 >> 8) & 0xFF;
        int b1 = (color1 >> 0) & 0xFF;

        int color2 = (int) Long.parseLong(nC2, 16);
        int r2 = (color2 >> 16) & 0xFF;
        int g2 = (color2 >> 8) & 0xFF;
        int b2 = (color2 >> 0) & 0xFF;

        double meanR = (r1 + r2) / 2;
        double meanG = (g1 + g2) / 2;
        double meanB = (b1 + b2) / 2;

        double distance = Math.sqrt(Math.pow((r1 - r2) / meanR, 2) + Math.pow((g1 - g2) / meanG, 2) + Math.pow((b1 - b2) / meanB, 2));

        if (distance < X) {
            return true;
        } else {
            return false;
        }

//        Color c1 = Color.decode("#"+nC1);
//        Color c2 = Color.decode("#"+nC2);
//        double meanR = (c1.getRed() + c2.getRed()) / 2;
//        double meanG = (c1.getGreen() + c2.getGreen()) / 2;
//        double meanB = (c1.getBlue() + c2.getBlue()) / 2;
//        double distance = Math.sqrt(Math.pow((c1.getRed() - c2.getRed())/meanR, 2) + Math.pow((c1.getGreen() - c2.getGreen())/meanG, 2) + Math.pow((c1.getBlue() - c2.getBlue())/meanB, 2));
//        if (distance < X) {
//            return true;
//        } else {
//            return false;
//        }
    }

    public static String shorterVersion(String toString) {
        if (toString.length() > 15) {
            return toString.substring(0, 15);
        }
        return toString;
    }

}

class FooterResume {

    private String type, text;
    private boolean isNumbered;

    public FooterResume(String type, String text, boolean isNumbered) {
        this.type = type;
        this.text = text;
        this.isNumbered = isNumbered;
    }

    public boolean isNumbered() {
        return this.isNumbered;
    }

    public String getType() {
        return this.type;
    }

    public String getText() {
        return this.text;
    }

    public boolean isDefault() {
        return getType().toLowerCase().equals("default");
    }

    public boolean isEven() {
        return getType().toLowerCase().equals("even");
    }

    public boolean isFirst() {
        return getType().toLowerCase().equals("first");
    }

    public String toString() {
        return isNumbered() + " " + getType() + " " + getText();
    }
}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

import java.util.Properties;
import java.io.InputStream;
import java.io.IOException;
import java.io.FileInputStream;

/**
 *
 * @author aavendan
 */
public class Styler {

    public static String headingNames[];
    public static int MAX_HEADINGS = 2;
    private static Properties heading[];

    public Styler() {
        FileInputStream input[] = new FileInputStream[MAX_HEADINGS];
        heading = new Properties[MAX_HEADINGS];
        headingNames = new String[MAX_HEADINGS];

        try {

            for (int i = 0; i < MAX_HEADINGS; i++) {
                heading[i] = new Properties();
                input[i] = new FileInputStream("heading" + (i + 1) + ".properties");
                heading[i].load(input[i]);
                headingNames[i] = "Heading" + String.valueOf(i + 1);

            }

        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {

            for (int i = 0; i < MAX_HEADINGS; i++) {
                if (input[i] != null) {
                    try {
                        input[i].close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

        }
    }

    public String getHeadingProperty(int id, String key) {
        return this.heading[id].getProperty(key);
    }

    public int getIndex(String name) {
        for (int i = 0; i < this.MAX_HEADINGS; i++) {
            if (headingNames[i].compareTo(name) == 0) {
                return i;
            }
        }
        return -1;
    }

}
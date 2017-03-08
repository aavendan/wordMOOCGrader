/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

import java.util.*;

/**
 *
 * @author aavendan
 */
public class NewMain {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here

        ArrayList list1 = new ArrayList();
        ArrayList list2 = new ArrayList();

        list1.add("a");
        list1.add("b");
        list1.add("c");
        list1.add("d");

        list2.add("a");
        list2.add("g");
        list2.add("c");
        list2.add("m");

        list1.stream().forEach(e1 -> list2.stream().forEach(e2 -> {
            if (e1.toString().compareTo(e2.toString()) == 0) {
                System.out.println(e1 + " " + e2);
            }
        }));

    }

}

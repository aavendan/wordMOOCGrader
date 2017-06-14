/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

/**
 *
 * @author Saul
 */
class StringSimilarity {
 
    /**
     * Calculates the similarity (a number within 0 and 1) between two strings.
     */
    public static double similarity(String s1, String s2) {
        String longer = s1, shorter = s2;
        if (s1.length() < s2.length()) { // longer should always have greater length
            longer = s2; shorter = s1;
        }
        int longerLength = longer.length();
        if (longerLength == 0) { return 1.0;}
        return (longerLength - editDistance(longer, shorter)) / (double) longerLength;
 
    }
 
    public static int editDistance(String s1, String s2) {
        s1 = s1.toLowerCase();
        s2 = s2.toLowerCase();
 
        int[] costs = new int[s2.length() + 1];
        for (int i = 0; i <= s1.length(); i++) {
            int lastValue = i;
            for (int j = 0; j <= s2.length(); j++) {
                if (i == 0)
                    costs[j] = j;
                else {
                    if (j > 0) {
                        int newValue = costs[j - 1];
                        if (s1.charAt(i - 1) != s2.charAt(j - 1))
                            newValue = Math.min(Math.min(newValue, lastValue),
                                    costs[j]) + 1;
                        costs[j - 1] = lastValue;
                        lastValue = newValue;
                    }
                }
            }
            if (i > 0)
                costs[s2.length()] = lastValue;
        }
        return costs[s2.length()];
    }
 
    public static void printSimilarity(String s, String t) {
        System.out.println(String.format(
            "%.3f is the similarity between \"%s\" and \"%s\"", similarity(s, t), s, t));
    }
 
    public static void main(String[] args) {
        printSimilarity("", "");
        printSimilarity("1234567890", "1");
        printSimilarity("1234567890", "123");
        printSimilarity("1234567890", "1234567");
        printSimilarity("1234567890", "1234567890");
        printSimilarity("1234567890", "1234567980");
        printSimilarity("47/2010", "472010");
        printSimilarity("47/2010", "472011");
        printSimilarity("47/2010", "AB.CDEF");
        printSimilarity("47/2010", "4B.CDEFG");
        printSimilarity("47/2010", "AB.CDEFG");
        printSimilarity("The quick fox jumped", "The fox jumped");
        printSimilarity("The quick fox jumped", "The fox");
        printSimilarity("n espectáculo aparte es su salto de mayor caudal y, con 80 m, también el más alto: la Garganta del diablo, el cual se puede disfrutar en toda su majestuosidad desde solo 50 m, recorriendo las pasarelas que parten desde Puerto Canoas, al que se llega utilizando el servicio de trenes ecológicos.", "Un espectáculo aparte es su salto de mayor caudal y, con 80 m, también el más alto: la Garganta del diablo, el cual se puede disfrutar en toda su majestuosidad desde solo 50 m, recorriendo las pasarelas que parten desde Puerto Canoas, al que se llega utilizando el servicio de trenes.");
        printSimilarity("or este salto pasa la frontera entre ambos países.", "or este salta pasa la frontera entre ambos paíse");
    }
 
}
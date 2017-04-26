/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

import java.nio.file.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import javax.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTNonVisualDrawingProps;
import org.docx4j.dml.picture.Pic;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.FooterReference;
import org.docx4j.wml.Ftr;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.P;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.Style;

/**
 *
 * @author aavendan
 */
public class Verifier {

    private static final double TOC_THRESHOLD_STYLE = 0.70;
    private static final double TOC_THRESHOLD_HEADINGS = 0.70;
    private static final double TOC_THRESHOLD_ELEMENTSINTOC = 0.40;
    private static final double COLUMN_THRESHOLD_SAME_PARAGRAPHS = 0.70;
    private static final double BREAKS_THRESHOLD_SAME_PAGEBREAKS = 0.40;
    private static final double PARAGRAPH_THRESHOLD_SAME_STYLE = 0.50;
    //    private static final int NAME_POSITION_BDR = 1;

    //Evaluation limits
    private static final double CD_LIMIT_2 = 0.70;
    private static final double CD_LIMIT_1 = 0.40;

    private static final double FOOTNOTE_LIMIT_2 = 0.89;
    private static final double FOOTNOTE_LIMIT_1 = 0.39;

    private static final double COLUMN_LIMIT_2 = 0.89;
    private static final double COLUMN_LIMIT_1 = 0.39;

    private static final double BDR_LIMIT_4 = 0.10;
    private static final double BDR_LIMIT_3 = 0.39;
    private static final double BDR_LIMIT_2 = 0.65;
    private static final double BDR_LIMIT_1 = 0.89;

    private static final double FOOTER_LIMIT_2 = 0.89;
    private static final double FOOTER_LIMIT_1 = 0.39;

    private static final double BULLET_LIMIT_4 = 0.89;
    private static final double BULLET_LIMIT_3 = 0.65;
    private static final double BULLET_LIMIT_2 = 0.39;
    private static final double BULLET_LIMIT_1 = 0.10;

    private static final double BREAK_LIMIT_4 = 0.89;
    private static final double BREAK_LIMIT_3 = 0.65;
    private static final double BREAK_LIMIT_2 = 0.39;
    private static final double BREAK_LIMIT_1 = 0.10;

    private static final double DFORMAT_LIMIT_4 = 0.89;
    private static final double DFORMAT_LIMIT_3 = 0.65;
    private static final double DFORMAT_LIMIT_2 = 0.39;
    private static final double DFORMAT_LIMIT_1 = 0.10;

    private static final int MAX_ELEMENTS = 2;
    public static int INDEX_ORIGINAL = 0;
    public static int INDEX_RESPONSE = 1;

    private static int GRADE_COLUMNS = 10;
    private static int GRADE_BORDER = 10;
    private static int GRADE_FOOTNOTE = 7;
    private static int GRADE_CAP = 6;
    private static int GRADE_TOC = 25;
    private static int GRADE_FOOTER = 10;
    private static int GRADE_BULLET = 7;
    private static int GRADE_BREAK = 10;
    private static int GRADE_DFORMAT = 15;
    private static int GRADE_TOTAL = 100;

    private static int FOOTER_FIRST = 0;
    private static int FOOTER_DEFAULT = 1;
    private static int FOOTER_EVEN = 2;

    public static String MULTIPLE = "Múltiple";
    public static String SINGLE_SPACING = "240";
    public static String MULTIPLE_SPACING = "720";

    private static WordprocessingMLPackage wordMLPackage[];
    private final String fileName[];
    private final Styler styler;

    public int totalGrade;
    public LinkedList grades;
    public java.io.FileWriter fwIndividual, fwGrupal;
    public String header = "";
    public boolean firstTime = false;

    public Verifier() {
        totalGrade = 0;
        grades = new LinkedList();
        wordMLPackage = new WordprocessingMLPackage[Verifier.MAX_ELEMENTS];
        fileName = new String[Verifier.MAX_ELEMENTS];
        styler = new Styler();
    }

    public void setFileName(int index, String fileName) {
        this.fileName[index] = fileName;
    }

    public String getFileName(int index) {
        return this.fileName[index];
    }

    private void loadDocument(int index) throws Exception {
        wordMLPackage[index] = WordprocessingMLPackage.load(new java.io.File(this.getFileName(index)));
    }

    public LinkedList getStyleObjectByQuery(int index, String query) throws Exception {
        List<Object> rList = wordMLPackage[index].getMainDocumentPart().getStyleDefinitionsPart().getJAXBNodesViaXPath(query, false);
        LinkedList objs = new LinkedList();

        rList.stream().forEach((jaxbNode) -> {
            objs.add(jaxbNode);
        });

        return objs;
    }

    public static LinkedList getDocumentObjectByQuery(int index, String query) throws Exception {
        List<Object> rList = wordMLPackage[index].getMainDocumentPart().getJAXBNodesViaXPath(query, false);
        LinkedList objs = new LinkedList();

        rList.stream().forEach((jaxbNode) -> {
            objs.add(jaxbNode);
        });

        return objs;
    }

    /*OLD by sections*/
    private LinkedList<FooterPart> getFooter(int index, int type) throws Exception {
        List<SectionWrapper> rList = wordMLPackage[index].getDocumentModel().getSections();
        LinkedList<FooterPart> objs = new LinkedList();

        rList.stream().forEach((sectPr) -> {
            if (sectPr.getSectPr().getType() == null) {
                HeaderFooterPolicy hfp = sectPr.getHeaderFooterPolicy();
                if (type == Verifier.FOOTER_FIRST) {
                    objs.add(hfp.getFirstFooter());
                }
                if (type == Verifier.FOOTER_DEFAULT) {
                    objs.add(hfp.getDefaultFooter());
                }
                if (type == Verifier.FOOTER_EVEN) {
                    objs.add(hfp.getEvenFooter());
                }
            }
        });

        return objs;
    }

    public JaxbXmlPart<org.docx4j.wml.Ftr> getFooterByQuery(int index, String id) throws Exception {

        RelationshipsPart rp2 = wordMLPackage[index].getMainDocumentPart().getRelationshipsPart();
        Relationship rel = rp2.getRelationshipByID(id);
        return (JaxbXmlPart) rp2.getPart(rel);

//        JaxbXmlPart<org.docx4j.wml.Ftr> jpart = (JaxbXmlPart) rp2.getPart(rel);
//        Ftr ftr = jpart.getContents();
//        writeReport(id);
//        writeReport(jpart.getXML());
//        writeReport(Helper.getTextFromP(jpart.getContents().getContent()));
//
//        return null;
    }

    public LinkedList getFootnotesObjectByQuery(int index, String query) throws Exception {
        List<Object> rList = wordMLPackage[index].getMainDocumentPart().getFootnotesPart().getJAXBNodesViaXPath(query, false);
        LinkedList objs = new LinkedList();

        rList.stream().forEach((jaxbNode) -> {
            objs.add(jaxbNode);
        });

        return objs;
    }

    private LinkedList loadTOC(int index) throws Exception {
        List<Object> tocs = wordMLPackage[index].getMainDocumentPart().getJAXBNodesViaXPath("//w:hyperlink[contains(@w:anchor,'_Toc')]", false);
        LinkedList tocElements = new LinkedList();
        tocs.stream().forEach((jaxbNode) -> {
            tocElements.add((javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink>) jaxbNode);
        });
        return tocElements;
    }

    public LinkedList loadHeadings(int index) throws Exception {
        //Styles inherited of Heading's
//        List<Object> hHeading = wordMLPackage[index].getMainDocumentPart().getStyleDefinitionsPart().getJAXBNodesViaXPath("//w:style[w:basedOn[contains(@w:val,'Heading')] or w:name[contains(@w:val,'heading')]]", false);
        String query = "//w:style[w:basedOn[contains(@w:val,'Heading')] or w:name[contains(@w:val,'heading')]]";
        List<Object> hHeading = getStyleObjectByQuery(index, query);
        List<String> styleNames = new LinkedList();

        hHeading.stream().forEach((jaxbNode) -> {
            styleNames.add("contains(@w:val,\'" + ((org.docx4j.wml.Style) jaxbNode).getStyleId() + "\')");
        });
        String strStyleNames = String.join(" or ", styleNames);
        query = "//w:body/w:p[w:pPr[w:pStyle[" + strStyleNames + "]]]";

        LinkedList headings = new LinkedList();
        if (hHeading.size() > 0) {
            List<Object> strHeading = wordMLPackage[index].getMainDocumentPart().getJAXBNodesViaXPath(query, false);

            strHeading.stream().forEach((jaxbNode) -> {
                headings.add((org.docx4j.wml.P) jaxbNode);
            });
        }

        return headings;
    }

    public Style getStyleByStyleId(int index, Object o) throws Exception {

        P p = ((org.docx4j.wml.P) o);
        if (p.getPPr() == null || p.getPPr().getPStyle() == null) {
            return null;
        }

        String styleId = p.getPPr().getPStyle().getVal();
        String query = "//w:style[contains(@w:styleId,'" + styleId + "')]";
        List<Object> objects = getStyleObjectByQuery(index, query);
        for (Object jaxbNode : objects) {
            return (org.docx4j.wml.Style) jaxbNode;
        }
        return null;
    }

    public int matchStyle(String query1, String query2, String query3, String queryd, String querys, String queryb) throws Exception {
        LinkedList rDoc, rStyleId, rBasedOn;

        rDoc = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, queryd);
        if (rDoc.size() > 0) {
            rDoc = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query1);
            if (rDoc.size() > 0) {
                //System.out.print("\t\tIn doc\n");
                return 1;
            }
            return 0;
        } else {
            rStyleId = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, querys);
            if (rStyleId.size() > 0) {
                rStyleId = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, query2);
                if (rStyleId.size() > 0) {
                    //System.out.print("\t\tIn Style\n");
                    return 1;
                }
                return 0;
            } else {
                rBasedOn = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, queryb);
                if (rBasedOn.size() > 0) {
                    rBasedOn = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, query3);
                    if (rBasedOn.size() > 0) {
                        //System.out.print("\t\tIn BasedOn\n");
                        return 1;
                    }
                }
                //System.out.print("\t\tNONE\n");
                return 0;
            }
        }
    }

    private boolean hasListing(P p, Style s) {
        if (p.getPPr() != null && p.getPPr().getNumPr() != null) {
            return true;
        }
        if (s.getPPr() != null && s.getPPr().getNumPr() != null) {
            return true;
        }
        return false;
    }

    private boolean sameHeadingName(Style sOriginal, Style sResponse) throws Exception {
        return sOriginal.getStyleId().compareTo(sResponse.getStyleId()) == 0
                || sOriginal.getStyleId().compareTo(sResponse.getBasedOn().getVal()) == 0
                || sOriginal.getBasedOn().getVal().compareTo(sResponse.getStyleId()) == 0
                || sOriginal.getBasedOn().getVal().compareTo(sResponse.getBasedOn().getVal()) == 0;
    }

    private String getHeadingName(Style sOriginal, Style sResponse) throws Exception {
        if (sOriginal.getBasedOn().getVal().compareTo(sResponse.getBasedOn().getVal()) == 0 || sOriginal.getStyleId().compareTo(sResponse.getBasedOn().getVal()) == 0) {
            return sResponse.getBasedOn().getVal();
        }
        if (sOriginal.getBasedOn().getVal().compareTo(sResponse.getStyleId()) == 0) {
            return sResponse.getStyleId();
        }
        return sResponse.getStyleId();
    }

    public static boolean isMultiple(String line_spacing) {
        return line_spacing.compareTo(MULTIPLE) == 0;
    }

    private boolean checkStyleHeading(Style sOriginal, Style sResponse, P pResponse) throws Exception {
        String headingName = getHeadingName(sOriginal, sResponse);
        int indexHeading = styler.getIndex(headingName);

        if (sameHeadingName(sOriginal, sResponse)) {

            String fontname, size, bold, hexcolor, spacing_before, spacing_after, line_spacing;
            String queryd, querys, queryb, query1, query2, query3;
            int values, check3, check4, check5, check6, check7, check8, check9;

            fontname = styler.getHeadingProperty(indexHeading, "fontname");
            size = String.valueOf(Integer.valueOf(styler.getHeadingProperty(indexHeading, "size")) * 2);
            bold = styler.getHeadingProperty(indexHeading, "bold");
            hexcolor = styler.getHeadingProperty(indexHeading, "hexcolor");
            spacing_before = String.valueOf(Integer.valueOf(styler.getHeadingProperty(indexHeading, "spacing_before")) * 20);
            spacing_after = String.valueOf(Integer.valueOf(styler.getHeadingProperty(indexHeading, "spacing_after")) * 20);
            line_spacing = styler.getHeadingProperty(indexHeading, "line_spacing");
            values = Integer.valueOf(styler.getHeadingProperty(indexHeading, "values"));

            //Check in document' style and style part
            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:rFonts[@w:cs and string-length(@w:cs)!=0]]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:rFonts[@w:ascii and string-length(@w:ascii)!=0]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:rFonts[@w:ascii and string-length(@w:ascii)!=0]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:rFonts[contains(@w:cs,'" + fontname + "')]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:rFonts[contains(@w:ascii,'" + fontname + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:rFonts[contains(@w:ascii,'" + fontname + "')]]]";
            //System.out.print("\t\tFont Name:");
            check3 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[contains(@w:val," + size + ")]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
            //System.out.print("\t\tFont Size:");
            check4 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:b[not(@*) or w:val='true']]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:b[not(@*) or w:val='true']]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:b[not(@*) or w:val='true']]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:b[not(@*)]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:b[not(@*)]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:b[not(@*)]]]";
            //System.out.print("\t\tBold:");
            check5 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:color[@w:val]]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:color[@w:val]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:color[@w:val]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]";
            //System.out.print("\t\tColor:");
            check6 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:before]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[@w:before]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[@w:before]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
            //System.out.print("\t\tSpacing Before:");
            check7 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:after]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[@w:after]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[@w:after]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
            //System.out.print("\t\tSpacing After:");
            check8 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            String value_line_spacing = Verifier.SINGLE_SPACING;
            if (Verifier.isMultiple(line_spacing)) {
                value_line_spacing = MULTIPLE_SPACING;
            }

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:line]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[@w:line]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[@w:line]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
            //System.out.print("\t\tLine Spacing:");
            check9 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            /*writeReport("Check3: " + check3);
             writeReport("Check4: " + check4);
             writeReport("Check5: " + check5);
             writeReport("Check6: " + check6);
             writeReport("Check7: " + check7);
             writeReport("Check8: " + check8);
             writeReport("Check9: " + check9);*/
            return (double) (check3 + check4 + check5 + check6 + check7 + check8 + check9) / values >= Verifier.TOC_THRESHOLD_STYLE;
        }

        return false;
    }

    public void validateTOC() throws Exception {

        int grade = 0;

        LinkedList tocResponse = loadTOC(Verifier.INDEX_RESPONSE);
        LinkedList tocOriginal = loadTOC(Verifier.INDEX_ORIGINAL);

        /*tocOriginal.stream().forEach(obj -> {
         javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink> h = (javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink>)obj;
         writeReport(Helper.getTextFromP(h.getValue().getContent()));
         });
         tocResponse.stream().forEach(obj -> {
         javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink> h = (javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink>)obj;
         writeReport(Helper.getTextFromP(h.getValue().getContent()));
         });*/
        LinkedList headingsOriginal = loadHeadings(Verifier.INDEX_ORIGINAL);
        LinkedList headingsResponse = loadHeadings(Verifier.INDEX_RESPONSE);

        /*headingsOriginal.stream().forEach(obj -> {
         P p = (P) obj;
         writeReport(Helper.getTextFromP(p.getContent()));
         });
         headingsResponse.stream().forEach(obj -> {
         P p = (P) obj;
         writeReport(Helper.getTextFromP(p.getContent()));
         });*/
        String tocRElement, tocOElement;
        int sameInOriginal = 0, notHere = 0, missing = 0, totalTOC;

        for (Iterator it = tocOriginal.iterator(); it.hasNext();) {
            Object oo = it.next();
            tocOElement = Helper.getTextFromP(((javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink>) oo).getValue().getContent());
            tocOElement = tocOElement.split("PAGEREF")[0].trim().toLowerCase();

            for (Iterator it2 = tocResponse.iterator(); it2.hasNext();) {
                Object or = it2.next();
                tocRElement = Helper.getTextFromP(((javax.xml.bind.JAXBElement<org.docx4j.wml.P.Hyperlink>) or).getValue().getContent());
                tocRElement = tocRElement.split("PAGEREF")[0].trim().toLowerCase();

                if (tocRElement.compareTo(tocOElement) == 0) {
                    sameInOriginal++;
                }
            }

        }

        //Styles: Same headingsOriginal on sResponse file
        P pResponse, pOriginal;
        String styleName, strOriginal, strResponse;
        Style sOriginal, sResponse;
        int foundStyle = 0, foundListing = 0;
        boolean exists, sameStyle, hasListing;

        for (Iterator it = headingsOriginal.iterator(); it.hasNext();) {
            Object o = it.next();
            strOriginal = Helper.getTextFromP(((org.docx4j.wml.P) o).getContent());

            exists = false;
            sameStyle = false;
            hasListing = false;

            for (Iterator it2 = headingsResponse.iterator(); it2.hasNext();) {
                Object o2 = it2.next();
                strResponse = Helper.getTextFromP(((org.docx4j.wml.P) o2).getContent());

                if (strOriginal.toLowerCase().trim().compareTo(strResponse.toLowerCase().trim()) == 0) {

                    exists = true;
                    pOriginal = (org.docx4j.wml.P) o;
                    pResponse = (org.docx4j.wml.P) o2;
                    sOriginal = getStyleByStyleId(Verifier.INDEX_ORIGINAL, o);
                    sResponse = getStyleByStyleId(Verifier.INDEX_RESPONSE, o2);
                    sameStyle = checkStyleHeading(sOriginal, sResponse, pResponse);
                    //writeReport("SameStyle: "+sameStyle);
                    hasListing = hasListing(pOriginal, sOriginal) && hasListing(pResponse, sResponse);
                    //writeReport("HasListingResponse: "+hasListing(pResponse, sResponse));
                    //writeReport("HasListing: "+hasListing);
                    break;

                }
            }

            if (sameStyle) {
                foundStyle++;
            }
            if (hasListing) {
                foundListing++;
            }

        }

        missing = tocOriginal.size() - sameInOriginal;
        notHere = tocResponse.size() - sameInOriginal;
        totalTOC = sameInOriginal - notHere - missing;

        writeReport("Grading: Table of Contents");

        //TOC: exist or not
        if (tocResponse.size() > 0) {
            grade += 5;
            writeReport("\tHas TOC +5");
            
            if ((double) totalTOC / tocOriginal.size() >= Verifier.TOC_THRESHOLD_ELEMENTSINTOC) {
                grade += 10;
                writeReport("\tMost elements in TOC! +10");
            } else {
                grade += 3;
                writeReport("\tFew elements in TOC +3");
            }

        } else {
            grade += 0;
            writeReport("\tWithout TOC +0");
        }

        //Correct Style
        if ((double) foundStyle / headingsOriginal.size() >= Verifier.TOC_THRESHOLD_HEADINGS) {
            grade += 5;
            writeReport("\tMost with correct styles! +5");
        } else {
            writeReport("\tFew headings with correct styles +0");
        }

        //Correct Listing
        if ((double) foundListing / headingsOriginal.size() >= Verifier.TOC_THRESHOLD_HEADINGS) {
            grade += 5;
            writeReport("\tMost with correct listing! +5");
        } else {
            writeReport("\tFew headings with correct listing +0");
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_TOC);
        grades.addFirst(grade);
        totalGrade += grade;
        header += "TOC;";
    }

    private boolean isFootnoteReference(Object o) {
        if (o instanceof org.docx4j.wml.R) {
            R c = (org.docx4j.wml.R) o;
            if (c.getRPr() != null && c.getRPr().getRStyle() != null && c.getRPr().getRStyle().getVal().compareTo("FootnoteReference") == 0) {
                return true;
            }
        }
        return false;
    }

    private String getFootNoteText(int index, R footnoteReference) throws Exception {
        List<Object> els = footnoteReference.getContent();

        for (Object el : els) {
            if (el instanceof javax.xml.bind.JAXBElement) {
                JAXBElement<org.docx4j.wml.CTFtnEdnRef> ds = (JAXBElement) el;
                org.docx4j.wml.CTFtnEdnRef val = ds.getValue();

                String query = "//w:footnote[@w:id='" + val.getId().toString() + "']/w:p/w:r/w:t";
                LinkedList texts = getFootnotesObjectByQuery(index, query);

                String value = "";
                for (Object c : texts) {
                    JAXBElement<org.docx4j.wml.Text> text = (JAXBElement<org.docx4j.wml.Text>) c;
                    if ((text.getValue().getValue().trim()).length() > 0) {
                        value = (text.getValue().getValue().trim());
                    }
                }
                return value;
            }
        }

        return null;
    }

    public HashMap<String, String> getFootNotes(int index, LinkedList footNoteInDoc) throws Exception {
        R footreference = null, previous = null;
        HashMap<String, String> footnotesMap = new HashMap();

        for (Object o1 : footNoteInDoc) {
            List<Object> listOfR = ((P) o1).getContent();

            for (Object o2 : listOfR) {
                if (o2 instanceof org.docx4j.wml.R) {
                    if (isFootnoteReference(o2)) {
                        String[] txtLst = Helper.getTextFromR(previous.getContent()).split(" ");
                        String originText = txtLst[txtLst.length - 1];
                        String footNotesText = getFootNoteText(index, (R) o2);
                        footnotesMap.put(originText.toLowerCase(), footNotesText.toLowerCase());
                    }
                    previous = (R) o2;
                }
            }
        }

        return footnotesMap;
    }

    public void validateFootNote() throws Exception {

        int grade = 0;

        String query;
        query = "//w:p[w:r[w:rPr[w:rStyle[@w:val='FootnoteReference']]]]";

        LinkedList footNoteOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList footNoteResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        HashMap<String, String> hashOriginal = getFootNotes(Verifier.INDEX_ORIGINAL, footNoteOriginal);
        HashMap<String, String> hashResponse = getFootNotes(Verifier.INDEX_RESPONSE, footNoteResponse);

        int countKeyOriginal = hashOriginal.keySet().size();
        int countKeyResponse = hashResponse.keySet().size();

        int count_specs = 0;
        double total_specs = 0;
        boolean sameWords = true, sameText = true;

        //has footnotes
        total_specs++;
        if (countKeyResponse > 0) {
            count_specs++;
        }

        //same number of footnotes
        total_specs++;
        if (countKeyOriginal == countKeyResponse) {
            count_specs++;
        } else {
            count_specs--;
        }

        for (String key1 : hashOriginal.keySet()) {
            sameWords = sameWords && hashResponse.containsKey(key1);
            sameText = hashResponse.containsKey(key1) ? sameText && (hashResponse.get(key1).compareTo(hashOriginal.get(key1)) == 0) : false;

        }

        total_specs++;
        if (sameWords) {
            count_specs++;
        }

        total_specs++;
        if (sameText) {
            count_specs++;
        }

        writeReport("Grading: Footnotes");

        if ((double) count_specs / total_specs > Verifier.FOOTNOTE_LIMIT_2) {
            grade += 7;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) count_specs / total_specs >= Verifier.FOOTNOTE_LIMIT_1) {
            grade += 4;
            writeReport("\t40-89% Specs! +" + grade);
        } else if ((double) count_specs / total_specs > 0 && (double) count_specs / total_specs < Verifier.FOOTNOTE_LIMIT_1) {
            grade += 1;
            writeReport("\t 0-39% Specs! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_FOOTNOTE);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Footnote;";
    }

    private void validateDropCap() throws Exception {

        int grade = 0;
        String query = "//w:p[w:pPr[w:framePr[@w:dropCap]]] | //w:p[w:pPr[w:framePr[@w:dropCap]]]/following-sibling::w:p[1]";
        LinkedList dropCapOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList dropCapResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        AtomicInteger capCounterOriginal = new AtomicInteger(0);

        dropCapOriginal.stream().forEach((obj) -> {
            P p = ((P) obj);
            if (p.getPPr().getFramePr() != null) {
                capCounterOriginal.getAndIncrement();
            }
        });

        P rOp, rOv, pOp, pOv;
        String capR, txtR, capO, txtO;
        int counterDL = 0, counterD = 0, counterML = 0, counterM = 0;

        for (Iterator it1 = dropCapResponse.iterator(); it1.hasNext();) {
            rOp = (P) it1.next();
            capR = Helper.getTextFromP(((P) rOp).getContent()).trim();
            rOv = (P) it1.next();
            //txtR = Helper.getTextFromP(((P) rOv).getContent()).substring(0, 20).trim();
            txtR = Helper.shorterVersion(Helper.getTextFromP(((P) rOv).getContent()).trim());

//            writeReport(rOp.getPPr().getFramePr().getDropCap().value()+" "+ capR+" "+txtR);
            for (Iterator it2 = dropCapOriginal.iterator(); it2.hasNext();) {
                pOp = (P) it2.next();
                capO = Helper.getTextFromP(((P) pOp).getContent()).trim();
                pOv = (P) it2.next();
                //txtO = Helper.getTextFromP(((P) pOv).getContent()).substring(0, 20).trim();
                txtO = Helper.shorterVersion(Helper.getTextFromP(((P) pOv).getContent()).trim());

                //match on text
                if (capR.compareTo(capO) == 0 && txtR.compareTo(txtO) == 0) {

//                    writeReport("\t"+pOp.getPPr().getFramePr().getDropCap().value()+" "+capO+" "+txtO);
                    if (rOp.getPPr().getFramePr().getDropCap().value().compareTo(pOp.getPPr().getFramePr().getDropCap().value()) == 0) {
                        if (rOp.getPPr().getFramePr().getLines().intValue() == pOp.getPPr().getFramePr().getLines().intValue()) {
                            counterDL++;
                        } else {
                            counterD++;
                        }
                    } else {

                        if (rOp.getPPr().getFramePr().getLines().intValue() == pOp.getPPr().getFramePr().getLines().intValue()) {
                            counterML++;
                        } else {
                            counterM++;
                        }
                    }
                }

            }
        }

        writeReport("Grading: Capital Letters");
//        writeReport((double) counterDL / capCounterOriginal.doubleValue() + " " + (double) counterD / capCounterOriginal.doubleValue() + " " + (double) counterML / capCounterOriginal.doubleValue() + " " + (double) counterM / capCounterOriginal.doubleValue());
        if ((double) counterDL / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_2) {
            grade += 6;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) counterD / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_1) {
            grade += 5;
            writeReport("\t66% - 89%! +" + grade);
        } else if ((double) counterD / capCounterOriginal.doubleValue() < Verifier.CD_LIMIT_1 && (double) counterML / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_1 && (double) counterML / capCounterOriginal.doubleValue() <= Verifier.CD_LIMIT_2) {
            grade += 4;
            writeReport("\t40% - 65%! +" + grade);
        } else if ((double) counterML / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_2 && (double) counterD / capCounterOriginal.doubleValue() < Verifier.CD_LIMIT_1) {
            grade += 2;
            writeReport("\t11% - 39%! +" + grade);
        } else if ((double) counterM / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_1 && (double) counterM / capCounterOriginal.doubleValue() <= Verifier.CD_LIMIT_2) {
            grade += 1;
            writeReport("\t 0% - 10%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_CAP);
        grades.addLast(grade);
        totalGrade += grade;
        header += "DropCap;";
    }

    private void validateColumns() throws Exception {

        int grade = 0;

        int specs = 0, totalSpecs = 0;
        String nquery, textO, textR;
        int countDifferentText;
        P pnO, pnR, pO1, pO2, pR1, pR2;

        String query = "//w:p[w:pPr[w:sectPr[w:cols[@w:num]]]] | //w:p[w:pPr[w:sectPr[w:cols[@w:num]]]]/preceding-sibling::w:p[w:pPr[w:sectPr[w:cols]]][1]";
        LinkedList dropCapOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList dropCapResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        //Has Columns
        totalSpecs++;
        if (dropCapResponse.size() > 0) {
            specs++;
        }

        //Same Groups of Columns
        totalSpecs++;
        specs = (dropCapOriginal.size() == dropCapResponse.size()) ? specs + 1 : specs - 1;

//        Igualdad(Grupo O,Grupo R) >= 70% -> #Parrafos, #ParrafosIguales, Separador, #Columnas
        Iterator itResponse;
        Iterator itOriginal = dropCapOriginal.iterator();

        while (itOriginal.hasNext()) {
            pO1 = (P) itOriginal.next();
            pO2 = (P) itOriginal.next();

            nquery = "//w:p[preceding-sibling::w:p[@w14:paraId='" + pO1.getParaId() + "'] and following-sibling::w:p[@w14:paraId='" + pO2.getParaId() + "']]";
            LinkedList linsideO = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, nquery);

            boolean found = false;

            itResponse = dropCapResponse.iterator();
            while (itResponse.hasNext()) {

                pR1 = (P) itResponse.next();
                pR2 = (P) itResponse.next();

                nquery = "//w:p[preceding-sibling::w:p[@w14:paraId='" + pR1.getParaId() + "'] and following-sibling::w:p[@w14:paraId='" + pR2.getParaId() + "']]";
                LinkedList linsideR = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, nquery);

                AtomicInteger similars = new AtomicInteger(0);

                //Does it match with a group text within columns?
                //Count similar columns
                linsideO.stream().forEach(p1 -> linsideR.stream().forEach(p2 -> {
                    P pp1 = (P) p1;
                    P pp2 = (P) p2;

                    String txt1 = Helper.getTextFromP(pp1.getContent()).toLowerCase().trim();
                    String txt2 = Helper.getTextFromP(pp2.getContent()).toLowerCase().trim();

                    if (txt1.compareTo(txt2) == 0) {
                        similars.getAndIncrement();
                    }

                }));

//                writeReport((double) similars.doubleValue() / linsideR.size());
                //Similar columns/total > threshold
                //If and only if they're similars paragraphs then check others specs
                if ((double) similars.doubleValue() / linsideR.size() >= Verifier.COLUMN_THRESHOLD_SAME_PARAGRAPHS) {

                    found = true;

                    //Same number of columns by columns group
                    totalSpecs++;
                    specs = (pO2.getPPr().getSectPr().getCols().getNum().intValue() == pR2.getPPr().getSectPr().getCols().getNum().intValue()) ? specs + 1 : specs - 1;
//                    writeReport((pO2.getPPr().getSectPr().getCols().getNum().intValue() == pR2.getPPr().getSectPr().getCols().getNum().intValue()) );

                    //Same number or paragraphs inside
                    totalSpecs++;
                    specs = (linsideO.size() == linsideR.size()) ? specs + 1 : specs - 1;
//                    writeReport(linsideO.size() == linsideR.size());

                    //Same separator by groups
                    totalSpecs++;
                    specs = (XmlUtils.marshaltoString(pO2, true, true).contains("w:sep") && XmlUtils.marshaltoString(pR2, true, true).contains("w:sep")) ? specs + 1 : specs - 1;
//                    writeReport(XmlUtils.marshaltoString(pO2, true, true).contains("w:sep") && XmlUtils.marshaltoString(pR2, true, true).contains("w:sep"));

                    totalSpecs++;
                    specs = (similars.intValue() == linsideR.size()) ? specs + 1 : specs - 1;
//                    writeReport("Same Texts: "+similars.intValue() +" "+ linsideR.size());
                }

                if (found) {
                    break;
                }

            }

            if (!found) {
                //There not exists a group of columns
                totalSpecs += 4;
            }

        }

        writeReport("Grading: Columns");
//        writeReport(specs+" "+totalSpecs);
        if ((double) specs / totalSpecs >= Verifier.COLUMN_LIMIT_2) {
            grade += 10;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.COLUMN_LIMIT_1) {
            grade += 5;
            writeReport("\t40% - 89%! +" + grade);
        } else {
            grade += 1;
            writeReport("\t 0% - 39%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_COLUMNS);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Columns;";
    }

    private void validateBdr() throws Exception {

        int grade = 0;

        String query = "//w:p[w:pPr[w:pBdr]]";
        LinkedList txtBdrOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList txtBdrResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        //Same number of borderO paragraphs
        double spec_numberPO = 0, spec_numberPR = 0;
        spec_numberPO = 1;
        if (txtBdrOriginal.size() == txtBdrResponse.size()) {
            spec_numberPR = 1;
        }
//        writeReport(spec_numberPR + " " + spec_numberPO);

        AtomicInteger i = new AtomicInteger(0);
        AtomicInteger countEqualTxtO = new AtomicInteger(0), countEqualTxtR = new AtomicInteger(0);
        AtomicInteger countSameBTypeO = new AtomicInteger(0), countSameBTypeR = new AtomicInteger(0);
        AtomicInteger countSameColorO = new AtomicInteger(0), countSameColorR = new AtomicInteger(0);
        AtomicInteger countSameShadingO = new AtomicInteger(0), countSameShadingR = new AtomicInteger(0);

        txtBdrOriginal.stream().forEach((objO) -> {
//        for (int i = 0; i < txtBdrOriginal.size(); i++) {
            i.getAndIncrement();

            P pO, pR;
            String txtO, txtR, colorR, colorO;
            CTBorder borderO, borderR;

            pO = (P) objO;//txtBdrOriginal.get(i);
            txtO = Helper.getTextFromP(pO.getContent());

            countEqualTxtO.getAndIncrement();

            for (int j = 0; j < txtBdrResponse.size(); j++) {
                pR = (P) txtBdrResponse.get(j);
                txtR = Helper.getTextFromP(pR.getContent());

                //Found
                //CHECK: if (txtO.trim().compareTo(txtR.trim()) == 0 || (i.intValue() == j && i.intValue() == Verifier.NAME_POSITION_BDR && txtR.length() > 0)) {
                if (txtO.trim().compareTo(txtR.trim()) == 0 && txtR.length() > 0) {
                    countEqualTxtR.getAndIncrement();

                    //Shading
                    if (pO.getPPr() != null && pO.getPPr().getShd() != null && pO.getPPr().getShd().getFill() != null) {
                        countSameShadingO.getAndIncrement();

                        if (pR.getPPr() != null && pR.getPPr().getShd() != null && pR.getPPr().getShd().getFill() != null) {

                            colorO = pO.getPPr().getShd().getFill();
                            colorR = pR.getPPr().getShd().getFill();
                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameShadingR.getAndIncrement();
                            }

                        }
                    }

                    //Border Top
                    borderO = pO.getPPr().getPBdr().getTop();
                    if (borderO != null) {

                        countSameColorO.getAndIncrement();
                        countSameBTypeO.getAndIncrement();
                        borderR = pR.getPPr().getPBdr().getTop();

                        if (borderR != null) {

//                            writeReport("Top");
//                            writeReport("O: " + borderO.getVal().value());
//                            writeReport("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                writeReport("T-Same " + colorO + " " + colorR);
                            } else {
//                                writeReport("T-Diff " + colorO + " " + colorR);
                            }
                        }
                    }

                    //Bottom
                    borderO = pO.getPPr().getPBdr().getBottom();
                    if (borderO != null) {

                        countSameColorO.getAndIncrement();
                        countSameBTypeO.getAndIncrement();
                        borderR = pR.getPPr().getPBdr().getBottom();

                        if (borderR != null) {

//                            writeReport("Bottom");
//                            writeReport("O: " + borderO.getVal().value());
//                            writeReport("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                writeReport("B-Same " + colorO + " " + colorR);
                            } else {
//                                writeReport("B-Diff " + colorO + " " + colorR);
                            }
                        }
                    }

                    //Left
                    borderO = pO.getPPr().getPBdr().getLeft();
                    if (borderO != null) {

                        countSameColorO.getAndIncrement();
                        countSameBTypeO.getAndIncrement();
                        borderR = pR.getPPr().getPBdr().getLeft();

                        if (borderR != null) {

//                            writeReport("Left");
//                            writeReport("O: " + borderO.getVal().value());
//                            writeReport("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                writeReport("L-Same " + colorO + " " + colorR);
                            } else {
//                                writeReport("L-Diff " + colorO + " " + colorR);
                            }
                        }
                    }

                    //Right
                    borderO = pO.getPPr().getPBdr().getRight();
                    if (borderO != null) {

                        countSameColorO.getAndIncrement();
                        countSameBTypeO.getAndIncrement();
                        borderR = pR.getPPr().getPBdr().getRight();

                        if (borderR != null) {

//                            writeReport("Right");
//                            writeReport("O: " + borderO.getVal().value());
//                            writeReport("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                writeReport("R-Same " + colorO + " " + colorR);
                            } else {
//                                writeReport("R-Diff " + colorO + " " + colorR);
                            }
                        }
                    }

                }

            }
        });

        //Same text in paragraphs
        double spec_samePO = 0, spec_samePR = 0;

        spec_samePO = 1;
        spec_samePR = (double) (countEqualTxtR.doubleValue() / (countEqualTxtO.doubleValue() == 0 ? 1 : countEqualTxtO.doubleValue()));
//        writeReport("Equal text: " + countEqualTxtR + " " + countEqualTxtO + " " + spec_samePR);

        //Same border in paragraphs
        double spec_sameBO = 0, spec_sameBR = 0;
        spec_sameBO = 1;
        spec_sameBR = (double) (countSameBTypeR.doubleValue() / (countSameBTypeO.doubleValue() == 0 ? 1 : countSameBTypeO.doubleValue()));
//        writeReport("Equal border: " + countSameBTypeR + " " + countSameBTypeO + " " + spec_sameBR);

        //Same border color in paragraphs
        double spec_sameCO = 0, spec_sameCR = 0;
        spec_sameCO = 1;
        spec_sameCR = (double) (countSameColorR.doubleValue() / (countSameColorO.doubleValue() == 0 ? 1 : countSameColorO.doubleValue()));
//        writeReport("Equal color: " + countSameColorR + " " + countSameColorO + " " + spec_sameCR);

        //Same shading in paragraphs
        double spec_sameSHO = 0, spec_sameSHR = 0;
        spec_sameSHO = 1;
        spec_sameSHR = (double) (countSameShadingR.doubleValue() / (countSameShadingO.doubleValue() == 0 ? 1 : countSameShadingO.doubleValue()));
//        writeReport("Equal shading: " + countSameShadingR + " " + countSameShadingO + " " + spec_sameSHR);

        //Images
        query = "//w:p[.//pic:spPr[.//a:prstDash or .//a:srgbClr]]//pic:pic";
        LinkedList imagesOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList imagesResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        double spec_numberIO = 0, spec_numberIR = 0;
        spec_numberIO = 1;
        if (imagesOriginal.size() == imagesResponse.size()) {
            spec_numberIR = 1;
        }
//        writeReport(spec_numberIR + " " + spec_numberIO);

        AtomicInteger sameNameImgO = new AtomicInteger(0), sameNameImgR = new AtomicInteger(0);
        AtomicInteger sameDashImgO = new AtomicInteger(0), sameDashImgR = new AtomicInteger(0);

//        imagesOriginal.stream().forEach((objO) -> {
        for (int k = 0; k < imagesOriginal.size(); k++) {
            Object objO = imagesOriginal.get(k);
            Pic picO = (Pic) ((javax.xml.bind.JAXBElement) objO).getValue(); //imagesOriginal.get(i)).getValue();
            String picNameO = picO.getNvPicPr().getCNvPr().getDescr();

            sameNameImgO.getAndIncrement();
            for (int j = 0; j < imagesResponse.size(); j++) {
                Pic picR = (Pic) ((javax.xml.bind.JAXBElement) imagesResponse.get(j)).getValue();
                String picNameR = picR.getNvPicPr().getCNvPr().getDescr();

                if (picNameO.compareTo(picNameR) == 0) {
                    sameNameImgR.getAndIncrement();

                    try {
                        String dashO = picO.getSpPr().getLn().getPrstDash().getVal().value();
                        String dashR = picR.getSpPr().getLn().getPrstDash().getVal().value();

                        sameDashImgO.getAndIncrement();
                        if (dashO.compareTo(dashR) == 0) {
                            sameDashImgR.getAndIncrement();
                        }
                    } catch (Exception ex) {
                    }

                }
            }
        }

        double spec_sameIO = 0, spec_sameIR = 0;

        spec_sameIO = 1;
        spec_sameIR = (double) (sameNameImgR.doubleValue() / (sameNameImgO.doubleValue() == 0 ? 1 : sameNameImgO.doubleValue()));
//        writeReport("Equal image: " + sameNameImgR + " " + sameNameImgO + " " + spec_sameIR);

        double spec_sameDashO = 0, spec_sameDashR = 0;

        spec_sameDashO = 1;
        spec_sameDashR = (double) (sameDashImgR.doubleValue() / (sameDashImgO.doubleValue() == 0 ? 1 : sameDashImgO.doubleValue()));
//        writeReport("Equal dash: " + sameDashImgR + " " + sameDashImgO + " " + spec_sameDashR);

        double totalO = spec_numberIO + spec_numberPO + spec_samePO + spec_sameIO + ((spec_sameDashO) / (spec_sameIO == 0 ? 1 : spec_sameIO)) + ((spec_sameBO + spec_sameCO + spec_sameSHO) / (spec_samePO == 0 ? 1 : spec_samePO));
        double totalR = spec_numberIR + spec_numberPR + spec_samePR + spec_sameIR + ((spec_sameDashR) / (spec_sameIR == 0 ? 1 : spec_sameIR)) + ((spec_sameBR + spec_sameCR + spec_sameSHR) / (spec_samePR == 0 ? 1 : spec_samePR));

        writeReport("Grading: Borders");

//        writeReport(totalR + " " + totalO);
        if ((double) totalR / totalO > Verifier.BDR_LIMIT_1) {
            grade += 10;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_2) {
            grade += 9;
            writeReport("\t66% - 89%! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_3) {
            grade += 7;
            writeReport("\t40% - 65%! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_4) {
            grade += 4;
            writeReport("\t11% - 39%! +" + grade);
        } else {
            grade += 1;
            writeReport("\t 0% - 10%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_BORDER);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Border;";
    }

    private boolean isFooterNumbered(FooterPart footer) {
        Iterator<Object> content = footer.getContent().iterator();
        while (content.hasNext()) {
            P val = (P) content.next();
            if (val.getPPr() != null && val.getPPr().getRPr() != null && val.getPPr().getRPr().getRStyle() != null) {
                return true;
            }
        }
        return false;
    }

    private String getContentFooter(FooterPart part) {
        String tmp, text = null;
        if (part != null) {

            List values = part.getContent();

            if (values != null) {
                Iterator it = values.iterator();

                text = new String();
                while (it.hasNext()) {
                    tmp = it.next().toString();
                    if (!tmp.contains("PAGE")) {
                        text += tmp;
                    }
                }
            }
        }
        return text;
    }

    private LinkedList<FooterResume> getFooters(int id) throws Exception {
        LinkedList<FooterResume> listResume = new LinkedList();

        String query = "//w:sectPr[w:footerReference]/w:footerReference";
        LinkedList lresponse = getDocumentObjectByQuery(id, query);

        FooterReference r;

        for (Object o1 : lresponse) {
            r = (org.docx4j.wml.FooterReference) o1;

            RelationshipsPart rp = wordMLPackage[id].getMainDocumentPart().getRelationshipsPart();

            List<Relationship> rels = rp.getRelationshipsByType(Namespaces.FOOTER);
            Iterator<Relationship> it = rels.iterator();

            while (it.hasNext()) {
                Relationship rel = it.next();
                JaxbXmlPart part = (JaxbXmlPart) rp.getPart(rel);
                if (rel.getId().equals(r.getId())) {
                    listResume.add(
                            new FooterResume(r.getType().value(),
                                    Helper.getTextFromFtr((Ftr) part.getContents()),
                                    Helper.isNumbered((Ftr) part.getContents())
                            )
                    );
                }
            }
        }

        return listResume;
    }

    public void validateFooter() throws Exception {

        int grade = 0;

        int specs = 0, totalSpecs = 0;
        boolean hasDefault = false, hasEven = false, hasFirst = false;

        LinkedList<FooterResume> lOriginal = getFooters(Verifier.INDEX_ORIGINAL);
        LinkedList<FooterResume> lResponse = getFooters(Verifier.INDEX_RESPONSE);

        FooterResume valO, valR;
        Iterator<FooterResume> itR, itO;

        itO = lOriginal.iterator();
        while (itO.hasNext()) {
            valO = itO.next();

            itR = lResponse.iterator();
            while (itR.hasNext()) {
                valR = itR.next();

                if (valO.getType().equals(valR.getType())) {

                    if (!hasDefault) {
                        hasDefault = valR.isDefault();

                        totalSpecs = valO.getText().length() > 0 ? totalSpecs + 1 : totalSpecs;
                        if (valR.getText().length() > 0 && valR.getText().equals(valO.getText())) {
                            specs++;
                        }

//                        writeReport("TEXT: "+totalSpecs + " " + valO + " " + specs + " " + valR);
                        totalSpecs++;
                        specs = valO.isNumbered() == valR.isNumbered() ? specs + 1 : specs;
//                        writeReport("NUMBERED: "+totalSpecs + " " + valO + " " + specs + " " + valR);

                    } else {
                    }

                    if (!hasFirst) {
                        hasFirst = valR.isFirst();

                        totalSpecs = valO.getText().length() > 0 ? totalSpecs + 1 : totalSpecs;
                        if (valR.getText().length() > 0 && valR.getText().equals(valO.getText())) {
                            specs++;
                        }

//                        writeReport("TEXT: "+totalSpecs + " " + valO + " " + specs + " " + valR);
                        totalSpecs++;
                        specs = valO.isNumbered() == valR.isNumbered() ? specs + 1 : specs;
//                        writeReport("NUMBERED: "+totalSpecs + " " + valO + " " + specs + " " + valR);

                    } else {
                    }

                    if (!hasEven) {
                        hasEven = valR.isEven();

                        totalSpecs = valO.getText().length() > 0 ? totalSpecs + 1 : totalSpecs;
                        if (valR.getText().length() > 0 && valR.getText().equals(valO.getText())) {
                            specs++;
                        }

//                        writeReport(totalSpecs + " " + valO + " " + specs + " " + valR);
                        totalSpecs++;
                        specs = valO.isNumbered() == valR.isNumbered() ? specs + 1 : specs;
//                        writeReport(totalSpecs + " " + valO + " " + specs + " " + valR);

                    } else {
                    }

                }
            }
        }

//        FooterPart elOriginal = null, elResponse = null;
//        HeaderFooterPolicy hfp;
//
//        Iterator<FooterPart> originalFooter, responseFooter;
//
//        int type = Verifier.FOOTER_DEFAULT;
//
//        originalFooter = getFooter(Verifier.INDEX_ORIGINAL, type).iterator();
//
//        while (originalFooter.hasNext()) {
//            elOriginal = originalFooter.next();
//
//            if (elOriginal != null) {
//
//                //Verify numbered footer
//                totalSpecs = isFooterNumbered(elOriginal) ? totalSpecs + 1 : totalSpecs;
//                totalSpecs = getContentFooter(elOriginal) != null ? totalSpecs + 1 : totalSpecs;
//                writeReport("O numbered: " + isFooterNumbered(elOriginal));
//                writeReport("O content: " + getContentFooter(elOriginal));
//
//            }
//
//            responseFooter = getFooter(Verifier.INDEX_RESPONSE, type).iterator();
//            while (responseFooter.hasNext()) {
//                elResponse = responseFooter.next();
//
//                if (elResponse != null) {
//                    specs = isFooterNumbered(elResponse) ? (isFooterNumbered(elOriginal) == isFooterNumbered(elResponse) ? specs + 1 : specs - 1) : specs;
//                    specs = getContentFooter(elResponse) != null ? (getContentFooter(elOriginal).equals(getContentFooter(elResponse)) ? specs + 1 : specs - 1) : specs;
//                    writeReport("R numbered: " + isFooterNumbered(elResponse));
//                    writeReport("R content: " + getContentFooter(elResponse));
//                }
//            }
//        }
        //Has footer
        totalSpecs++;
        if (hasDefault | hasEven | hasFirst) {
            specs++;
        }

        writeReport("Grading: Footer");

        if ((double) specs / totalSpecs >= Verifier.FOOTER_LIMIT_2) {
            grade += 10;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.FOOTER_LIMIT_1) {
            grade += 5;
            writeReport("\t40% - 89%! +" + grade);
        } else {
            grade += 1;
            writeReport("\t 0% - 39%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_FOOTER);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Footer;";
    }

    private LinkedList<P> getBullets(int id) throws Exception {
        LinkedList<P> pBullets = new LinkedList();

        String query = "//w:p[w:pPr[w:numPr[w:ilvl]]]";
        LinkedList lresponse = getDocumentObjectByQuery(id, query);

        for (Object o : lresponse) {
            pBullets.add((P) o);
        }

        return pBullets;
    }

    private ArrayList<String> getAbstract(Object contents, int numId, int ilvl) throws Exception {
        Numbering numbering = (org.docx4j.wml.Numbering) contents;
        ArrayList<String> values = new ArrayList();

        List<Numbering.Num> nums = numbering.getNum();

        for (Numbering.Num num : nums) {

            if (num.getNumId().intValue() == numId) {

                int abstractNumId = num.getAbstractNumId().getVal().intValue();
                List<Numbering.AbstractNum> numsAbstract = numbering.getAbstractNum();

                for (Numbering.AbstractNum abs : numsAbstract) {

                    if (abs.getAbstractNumId().intValue() == abstractNumId) {

                        List<Lvl> lvls = abs.getLvl();
                        for (Lvl lvl : lvls) {

                            if (lvl.getIlvl().intValue() == ilvl) {
                                values.add(Helper.escapeNonAscii(lvl.getLvlText().getVal()));
                                values.add(lvl.getNumFmt().getVal().value());
                                return values;
                            }

                        }

                    }

                }
            }
        }

        return null;
    }

    private ArrayList<String> getAbstractBulletLevel(int id, int numId, int ilvl) throws Exception {
        RelationshipsPart rp = wordMLPackage[id].getMainDocumentPart().getRelationshipsPart();
        ArrayList<String> values = null;

        List<Relationship> rels = rp.getRelationshipsByType(Namespaces.NUMBERING);
        Iterator<Relationship> it = rels.iterator();
        while (it.hasNext()) {
            Relationship rel = it.next();
            JaxbXmlPart part = (JaxbXmlPart) rp.getPart(rel);
            values = getAbstract(part.getContents(), numId, ilvl);
            if (values != null) {
                return values;
            }
        }

        return null;
    }

    public void validateBullet() throws Exception {

        int grade = 0;

        int specs = 0, totalSpecs = 0;
        int ilvlO = -1, ilvlR = -1, numIdO = -1, numIdR = -1;
        String contentO = "", contentR = "";
        ArrayList<String> valuesO, valuesR;

        LinkedList<P> pBulletOriginal = getBullets(Verifier.INDEX_ORIGINAL);
        LinkedList<P> pBulletResponse = getBullets(Verifier.INDEX_RESPONSE);

        //Has some
        totalSpecs = pBulletOriginal.size() > 0 ? totalSpecs + 1 : totalSpecs;
        specs = pBulletResponse.size() > 0 ? specs + 1 : specs;

        for (P p1 : pBulletOriginal) {

            //Counting its bullet type & symbol
            totalSpecs++;
            numIdO = p1.getPPr().getNumPr().getNumId().getVal().intValue();
            totalSpecs++;
            ilvlO = p1.getPPr().getNumPr().getIlvl().getVal().intValue();

            //Counting its text
            totalSpecs++;
            //contentO = Helper.getTextFromP(p1.getContent()).substring(0, 20);
            contentO = Helper.shorterVersion(Helper.getTextFromP(p1.getContent()));

            valuesO = getAbstractBulletLevel(Verifier.INDEX_ORIGINAL, numIdO, ilvlO);

//            writeReport(ilvlO + " " + numIdO + " " + contentO+" "+valuesO);
            for (P p2 : pBulletResponse) {

                numIdR = p2.getPPr().getNumPr().getNumId().getVal().intValue();
                ilvlR = p2.getPPr().getNumPr().getIlvl().getVal().intValue();

                //contentR = Helper.getTextFromP(p2.getContent()).substring(0, 19);
                contentR = Helper.shorterVersion(Helper.getTextFromP(p2.getContent()));

                valuesR = getAbstractBulletLevel(Verifier.INDEX_RESPONSE, numIdR, ilvlR);

                //Counting same text in response
                if (contentO.contains(contentR)) {
                    specs++;
//                    writeReport(ilvlR + " " + numIdR + " " + contentR);

                    if (valuesR != null && valuesO != null) {
                        //Counting same bullet type
                        if (valuesR.get(1).equals(valuesO.get(1))) {
                            specs++;

                            //Counting same bullet symbol
                            if (valuesR.get(0).equals(valuesO.get(0))) {
//                            writeReport(valuesR);
                                specs++;
                            }

                        }
                    }
                }

            }
        }

        writeReport("Grading: Bullets");

        if ((double) specs / totalSpecs >= Verifier.BULLET_LIMIT_4) {
            grade += 7;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BULLET_LIMIT_3) {
            grade += 6;
            writeReport("\t66% - 89%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BULLET_LIMIT_2) {
            grade += 5;
            writeReport("\t40% - 65%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BULLET_LIMIT_1) {
            grade += 3;
            writeReport("\t11% - 39%! +" + grade);
        } else {
            grade += 1;
            writeReport("\t 0% - 10%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_BULLET);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Bullets;";
    }

    public LinkedList getPageBreaks(int index) throws Exception {
        String query = "//w:p[w:r[w:br[contains(@w:type,'page')]]]";
        return getDocumentObjectByQuery(index, query);
    }

    public LinkedList getSectionBreaks(int index) throws Exception {
        String query = "//w:p[w:pPr[w:sectPr]]";
        return getDocumentObjectByQuery(index, query);
    }

    public String lookForImage(int index, P p) throws Exception {
        String query = "//w:p[@w14:paraId='" + p.getParaId() + "']//pic:cNvPr";
        LinkedList images = getDocumentObjectByQuery(index, query);
        if (images.size() > 0) {
            CTNonVisualDrawingProps props = (CTNonVisualDrawingProps) images.get(0);
            return props.getDescr();
        }

        return null;
    }

    public String lookForSdt(int index, P p) throws Exception {
        String text, query = "//w:sdt[w:sdtContent]//w:hyperlink//w:r";
        LinkedList ps = getDocumentObjectByQuery(index, query);
        ListIterator it2 = ps.listIterator(ps.size());

        while (it2.hasPrevious()) {
            Object o = it2.previous();
            if (o instanceof R) {
                text = Helper.getTextFromR(((R) o).getContent());
                if (text.length() > 0) {
                    return text;
                }
            }
        }

        return null;
    }

    public String getPreviousToSectionBreak(int index, P p, LinkedList elements) throws Exception {

        int i;
        for (i = 0; i < elements.size(); i++) {
            if (elements.get(i) instanceof P) {
                P p2 = (P) elements.get(i);
                if (p2.getParaId() != null && p2.getParaId().compareTo(p.getParaId()) == 0) {
                    break;
                }
            }
        }

        //looking for previousTextsO with any content
        for (int j = i - 1; j >= 0; j--) {
            //check for text
            if (elements.get(j) instanceof P && Helper.getTextFromP(((P) elements.get(j)).getContent()).length() > 0) {
                return Helper.getTextFromP(((P) elements.get(j)).getContent());
            }

            //check for image
            if (elements.get(j) instanceof P && lookForImage(index, (P) elements.get(j)) != null) {
                return lookForImage(index, (P) elements.get(j));
            }

            //check for sdt
            if (elements.get(j) instanceof P && lookForSdt(index, (P) elements.get(j)) != null) {
                return lookForSdt(index, (P) elements.get(j));
            }
        }

        return null;
    }

    public String getPreviousToBreak(int index, P p, LinkedList elements) throws Exception {
        String text = Helper.getTextFromP(p.getContent());

        //If there's some text then return itself
        if (text.length() > 0) {
            return text;
        }

        //Or an image
        text = lookForImage(index, p);
        if (text != null) {
            return text;
        }

        //Or sdt
        text = lookForSdt(index, p);
        if (text != null) {
            return text;
        }

        //Otherwise: look for something on previousTextsO
        //first: my position
        int i = elements.size();

        for (i = 0; i < elements.size(); i++) {
            if (elements.get(i) instanceof P) {
                P p2 = (P) elements.get(i);
                if (p2.getTextId() != null && p2.getTextId().compareTo(p.getTextId()) == 0) {
                    break;
                }
            }
        }

        //writeReport("RsidR " + p.getParaId() + " curent position: " + i);
        //looking for previousTextsO with any content
        for (int j = i - 1; j >= 0; j--) {
            //check for text
            if (elements.get(j) instanceof P && Helper.getTextFromP(((P) elements.get(j)).getContent()).length() > 0) {
                return Helper.getTextFromP(((P) elements.get(j)).getContent());
            }

            //check for image
            if (elements.get(j) instanceof P && lookForImage(index, (P) elements.get(j)) != null) {
                return lookForImage(index, (P) elements.get(j));
            }

            //check for sdt
            if (elements.get(j) instanceof P && lookForSdt(index, (P) elements.get(j)) != null) {
                return lookForSdt(index, (P) elements.get(j));
            }
        }

        return null;
    }

    public void validateBreaks() throws Exception {

        int grade = 0;

        int specs = 0, totalSpecs = 0;
        LinkedList elements, previousTextsO = new LinkedList(), previousTextsR = new LinkedList();
        Iterator breaks;
        P pBlock;

        String query = "//w:body/child::*";

        elements = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        breaks = getPageBreaks(Verifier.INDEX_ORIGINAL).iterator();
        while (breaks.hasNext()) {
            previousTextsO.add(getPreviousToBreak(Verifier.INDEX_ORIGINAL, (P) breaks.next(), elements));
        }

        elements = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);
        breaks = getPageBreaks(Verifier.INDEX_RESPONSE).iterator();
        while (breaks.hasNext()) {
            previousTextsR.add(getPreviousToBreak(Verifier.INDEX_RESPONSE, (P) breaks.next(), elements));
        }

        /*for(int i = 0; i < previousTextsO.size(); i++)
         writeReport("O: "+previousTextsO.get(i));
        
         for(int j = 0; j < previousTextsR.size(); j++)
         writeReport("R: "+previousTextsR.get(j));*/
        totalSpecs++;
        if (previousTextsO.size() == previousTextsR.size()) {
            specs++;
        }

        //boolean found;
        //Counting similars
        for (int i = 0; i < previousTextsO.size(); i++) {
            totalSpecs++;
            //found = false;
            for (int j = 0; j < previousTextsR.size(); j++) {
                if (previousTextsO.get(i) != null && previousTextsR.get(j) != null) {
                    if (Helper.shorterVersion(previousTextsO.get(i).toString()).compareTo(Helper.shorterVersion(previousTextsR.get(j).toString())) == 0) {
                        previousTextsR.remove(j);
                        specs++;
                        //found = true;
                    }
                }
            }
            /*if(!found)
             specs--; */
        }

        //writeReport(specs + " :: " + totalSpecs);
        previousTextsO.clear();
        elements = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        breaks = getSectionBreaks(Verifier.INDEX_ORIGINAL).iterator();
        while (breaks.hasNext()) {
            previousTextsO.add(getPreviousToSectionBreak(Verifier.INDEX_ORIGINAL, (P) breaks.next(), elements));
        }

        previousTextsR.clear();
        elements = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);
        breaks = getSectionBreaks(Verifier.INDEX_RESPONSE).iterator();
        while (breaks.hasNext()) {
            previousTextsR.add(getPreviousToSectionBreak(Verifier.INDEX_RESPONSE, (P) breaks.next(), elements));
        }

        /*for(int i = 0; i < previousTextsO.size(); i++)
         writeReport("O: "+previousTextsO.get(i));
        
         for(int j = 0; j < previousTextsR.size(); j++)
         writeReport("R: "+previousTextsR.get(j));*/
        totalSpecs++;
        if (previousTextsO.size() == previousTextsR.size()) {
            specs++;
        }

        int countSimilarities = 0, size = previousTextsO.size();
        totalSpecs++;
        for (int i = 0; i < previousTextsO.size(); i++) {
            //found = false;
            for (int j = 0; j < previousTextsR.size(); j++) {
                if (previousTextsO.get(i).toString().contains(Helper.shorterVersion(previousTextsR.get(j).toString()))) {
                    previousTextsR.remove(j);
                    countSimilarities++;
                    //found = true;
                }
            }
            /*if(!found)
             specs--;*/
        }

        specs = ((double) countSimilarities / size >= Verifier.BREAKS_THRESHOLD_SAME_PAGEBREAKS) ? specs + 1 : specs;

        //writeReport(specs+" :: "+totalSpecs);
        writeReport("Grading: Page Breaks and Sections");

        if ((double) specs / totalSpecs >= Verifier.BREAK_LIMIT_4) {
            grade += 10;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BREAK_LIMIT_3) {
            grade += 8;
            writeReport("\t66% - 89%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BREAK_LIMIT_2) {
            grade += 6;
            writeReport("\t40% - 65%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.BREAK_LIMIT_1) {
            grade += 4;
            writeReport("\t11% - 39%! +" + grade);
        } else {
            grade += 1;
            writeReport("\t 0% - 10%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_BREAK);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Break;";

    }

    public int chechStyleParagraph(P pResponse, Style sResponse) throws Exception {

        String fontname, size, spacing_before, spacing_after, line_spacing;
        String queryd, querys, queryb, query1, query2, query3;
        int values, check3, check4, check7, check8, check9;

        fontname = styler.getParagraphProperty("fontname");
        size = String.valueOf(Integer.valueOf(styler.getParagraphProperty("size")) * 2);
        spacing_before = String.valueOf(Integer.valueOf(styler.getParagraphProperty("spacing_before")) * 20);
        spacing_after = String.valueOf(Integer.valueOf(styler.getParagraphProperty("spacing_after")) * 20);
        line_spacing = styler.getParagraphProperty("line_spacing");
        values = Integer.valueOf(styler.getParagraphProperty("values"));

        //Check in document' style and style part
        queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:rFonts[@w:cs and string-length(@w:cs)!=0]]]]";
        querys = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:rPr[w:rFonts[@w:ascii and string-length(@w:ascii)!=0]]]";
        queryb = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:rPr[w:rFonts[@w:ascii and string-length(@w:ascii)!=0]]]";

        query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:rFonts[contains(@w:cs,'" + fontname + "')]]]]";
        query2 = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:rPr[w:rFonts[contains(@w:ascii,'" + fontname + "')]]]";
        query3 = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:rPr[w:rFonts[contains(@w:ascii,'" + fontname + "')]]]";
        //System.out.print("\t\tFont Name: "+fontname+"\n");
        check3 = matchStyle(query1, query2, query3, queryd, querys, queryb);

        queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]]";
        querys = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";
        queryb = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";

        query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[contains(@w:val," + size + ")]]]]";
        query2 = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
        query3 = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
        //System.out.print("\t\tFont Size: "+size+"\n");
        check4 = matchStyle(query1, query2, query3, queryd, querys, queryb);

        queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:before]]]";
        querys = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[@w:before]]]";
        queryb = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[@w:before]]]";

        query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
        query2 = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
        query3 = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
        //System.out.print("\t\tSpacing Before: "+spacing_before+"\n");
        check7 = matchStyle(query1, query2, query3, queryd, querys, queryb);

        queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:after]]]";
        querys = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[@w:after]]]";
        queryb = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[@w:after]]]";

        query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
        query2 = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
        query3 = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
        //System.out.print("\t\tSpacing After: "+spacing_after+"\n");
        check8 = matchStyle(query1, query2, query3, queryd, querys, queryb);

        String value_line_spacing = Verifier.SINGLE_SPACING;
        if (Verifier.isMultiple(line_spacing)) {
            value_line_spacing = MULTIPLE_SPACING;
        }

        queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:line]]]";
        querys = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[@w:line]]]";
        queryb = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[@w:line]]]";

        query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
        query2 = "//w:style[@w:styleId='" + ((sResponse != null) ? sResponse.getStyleId() : "") + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
        query3 = "//w:style[@w:styleId='" + ((sResponse != null && sResponse.getBasedOn() != null) ? sResponse.getBasedOn().getVal() : "") + "' and w:pPr[w:spacing[contains(@w:line,'" + value_line_spacing + "')]]]";
        //System.out.print("\t\tLine Spacing:");
        check9 = matchStyle(query1, query2, query3, queryd, querys, queryb);

        return ((double) (check3 + check4 + check7 + check8 + check9) / values) >= Verifier.PARAGRAPH_THRESHOLD_SAME_STYLE ? 1 : 0;
    }

    public void validateFormat() throws Exception {

        int grade = 0;

        LinkedList elementsR;
        P pElementR;
        Style sResponse;

        String query = "//w:p[.//w:t and not(.//w:hyperlink) and not(.//w:bookmarkStart) and not(ancestor::w:sdtContent) and not(.//w:pPr[w:framePr[@w:dropCap]])]";
        elementsR = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        int specs = 0, totalSpecs = 0;

        // 0 - Title
        // 1 - Name
        for (int i = 2; i < elementsR.size(); i++) {
            pElementR = (P) elementsR.get(i);
            sResponse = getStyleByStyleId(Verifier.INDEX_RESPONSE, pElementR);

            totalSpecs++;
            specs += chechStyleParagraph(pElementR, sResponse);
        }

        //writeReport(specs+" :: "+totalSpecs);
        writeReport("Grading: Document format");

        if ((double) specs / totalSpecs >= Verifier.DFORMAT_LIMIT_4) {
            grade += 15;
            writeReport("\tMost Specs! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.DFORMAT_LIMIT_3) {
            grade += 13;
            writeReport("\t66% - 89%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.DFORMAT_LIMIT_2) {
            grade += 10;
            writeReport("\t40% - 65%! +" + grade);
        } else if ((double) specs / totalSpecs >= Verifier.DFORMAT_LIMIT_1) {
            grade += 6;
            writeReport("\t11% - 39%! +" + grade);
        } else {
            grade += 2;
            writeReport("\t 0% - 10%! +" + grade);
        }

        writeReport("\tGrade: " + grade + "/" + Verifier.GRADE_DFORMAT);
        grades.addLast(grade);
        totalGrade += grade;
        header += "Format;";
    }

    public void validate() throws Exception {
        loadDocument(Verifier.INDEX_ORIGINAL);
        loadDocument(Verifier.INDEX_RESPONSE);

        createReport();
        validateTOC();
        validateFormat();
        validateBdr();
        validateFootNote();
        validateFooter();
        validateDropCap();
        validateColumns();
        validateBreaks();
        validateBullet();
        writeReport("Total Grade: " + totalGrade + "/" + Verifier.GRADE_TOTAL);
        grades.addLast(totalGrade);
        closeReport();
    }

    private void createReport() throws Exception {
        fwIndividual = new java.io.FileWriter(this.getFileName(Verifier.INDEX_RESPONSE).replace(".docx", "-grade.txt"));
        String path = this.getFileName(Verifier.INDEX_ORIGINAL);
        path = path.substring(0, path.lastIndexOf("/") + 1);

        firstTime = Files.notExists(Paths.get(path + "resume-responses.csv"));
        fwGrupal = new java.io.FileWriter(path + "resume-responses.csv", true);
    }

    private void writeReport(String line) throws Exception {
        if (fwIndividual != null) {
            fwIndividual.write(line + "\n");
        }
    }

    private void closeReport() throws Exception {
        if (fwIndividual != null) {
            fwIndividual.close();
        }
        if (fwGrupal != null) {

            if (firstTime) {
                fwGrupal.write("Nombre;" + header + "Total" + "\n");
            }

            grades.addFirst(this.getFileName(Verifier.INDEX_RESPONSE).replace(".doc", ""));
            for (int i = 0; i < grades.size() - 1; i++) {
                fwGrupal.write(grades.get(i) + ";");
            }
            fwGrupal.write(grades.get(grades.size() - 1) + "\n");
            fwGrupal.close();
        }
    }

    private void saveFiles() throws Exception {
        java.io.FileWriter fw = new java.io.FileWriter("document.xml");
        fw.write(wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getXML());
        fw.close();
        fw = new java.io.FileWriter("style.xml");
        fw.write(wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getStyleDefinitionsPart().getXML());
        fw.close();
    }

    private void showStyle() throws Exception {
        System.out.println(wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getStyleDefinitionsPart().getXML());
    }

    private void showXML() throws Exception {
        System.out.println(wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getXML());
    }

    private void showFooterXML() throws Exception {

        //Relaciones intermedias con los otros archivos
        RelationshipsPart rp = wordMLPackage[Verifier.INDEX_ORIGINAL].getMainDocumentPart().getRelationshipsPart();

        List<Relationship> rels = rp.getRelationshipsByType(Namespaces.NUMBERING);
        Iterator<Relationship> it = rels.iterator();
        while (it.hasNext()) {
            Relationship rel = it.next();
            JaxbXmlPart part = (JaxbXmlPart) rp.getPart(rel);
            writeReport(part.getXML());
        }
    }

}

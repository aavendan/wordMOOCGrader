/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mooc.grader;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import org.apache.xml.utils.XMLStringDefault;
import org.docx4j.XmlUtils;
import org.docx4j.dml.picture.Pic;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.EndnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.org.apache.poi.hpsf.SummaryInformation;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.FooterReference;
import org.docx4j.wml.Ftr;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;

/**
 *
 * @author aavendan
 */
public class Verifier {

    private static final double TOC_LIMIT_STYLE = 0.70;
    private static final double TOC_LIMIT_HEADINGS = 0.70;
    private static final double TOC_LIMIT_ELEMENTSINTOC = 0.40;

    private static final double CD_LIMIT_ALL = 0.70;
    private static final double CD_LIMIT_SOME = 0.40;

    private static final double FOOTNOTE_LIMIT_1 = 0.39;
    private static final double FOOTNOTE_LIMIT_2 = 0.89;

    private static final double COLUMN_LIMIT_1 = 0.39;
    private static final double COLUMN_LIMIT_2 = 0.89;
    private static final double COLUMN_MINIMUM_SAME_PARAGRAPHS = 0.70;

    private static final double BDR_LIMIT_4 = 0.10;
    private static final double BDR_LIMIT_3 = 0.39;
    private static final double BDR_LIMIT_2 = 0.65;
    private static final double BDR_LIMIT_1 = 0.89;

    private static final int NAME_POSITION_BDR = 1;

    private static final int MAX_ELEMENTS = 2;
    public static int INDEX_ORIGINAL = 0;
    public static int INDEX_RESPONSE = 1;

    private static int GRADE_COLUMNS = 10;
    private static int GRADE_BORDER = 10;
    private static int GRADE_FOOTNOTE = 7;
    private static int GRADE_CAP = 6;
    private static int GRADE_TOC = 25;

    private static int FOOTER_FIRST = 0;
    private static int FOOTER_DEFAULT = 1;
    private static int FOOTER_EVEN = 2;

    private final WordprocessingMLPackage wordMLPackage[];
    private final String fileName[];
    private final Styler styler;

    public Verifier() {
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

    public LinkedList getDocumentObjectByQuery(int index, String query) throws Exception {
        List<Object> rList = wordMLPackage[index].getMainDocumentPart().getJAXBNodesViaXPath(query, false);
        LinkedList objs = new LinkedList();

        rList.stream().forEach((jaxbNode) -> {
            objs.add(jaxbNode);
        });

        return objs;
    }

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
//        System.out.println(id);
//        System.out.println(jpart.getXML());
//        System.out.println(Helper.getTextFromP(jpart.getContents().getContent()));
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
        List<Object> hHeading = wordMLPackage[index].getMainDocumentPart().getStyleDefinitionsPart().getJAXBNodesViaXPath("//w:style[w:basedOn[contains(@w:val,'Heading')] or w:name[contains(@w:val,'heading')]]", false);
        List<String> styleNames = new LinkedList();

        hHeading.stream().forEach((jaxbNode) -> {
            styleNames.add("contains(@w:val,\'" + ((org.docx4j.wml.Style) jaxbNode).getStyleId() + "\')");
        });
        String strStyleNames = String.join(" or ", styleNames);
        String query = "//w:body/w:p[w:pPr[w:pStyle[" + strStyleNames + "]]]";

        LinkedList headings = new LinkedList();
        if (hHeading.size() > 0) {
            List<Object> strHeading = wordMLPackage[index].getMainDocumentPart().getJAXBNodesViaXPath(query, false);

            strHeading.stream().forEach((jaxbNode) -> {
                headings.add((org.docx4j.wml.P) jaxbNode);
            });
        }

        return headings;
    }

    public Style getStyleByHeadingName(int index, Object o) throws Exception {
        String headingName = ((org.docx4j.wml.P) o).getPPr().getPStyle().getVal();
        List<Object> hHeading = wordMLPackage[index].getMainDocumentPart().getStyleDefinitionsPart().getJAXBNodesViaXPath("//w:style[contains(@w:styleId,'" + headingName + "')]", false);
        for (Object jaxbNode : hHeading) {
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
//                System.out.print("\t\tIn doc\n");
                return 1;
            }
            return 0;
        } else {
            rStyleId = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, querys);
            if (rStyleId.size() > 0) {
                rStyleId = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, query2);
                if (rStyleId.size() > 0) {
//                    System.out.print("\t\tIn Style\n");
                    return 1;
                }
                return 0;
            } else {
                rBasedOn = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, queryb);
                if (rBasedOn.size() > 0) {
                    rBasedOn = getStyleObjectByQuery(Verifier.INDEX_RESPONSE, query3);
                    if (rBasedOn.size() > 0) {
//                        System.out.print("\t\tIn BasedOn\n");
                        return 1;
                    }
                }
//                System.out.print("\t\tNONE\n");
                return 0;
            }
        }
    }

    private boolean checkStyle(Style sOriginal, Style sResponse, P pResponse) throws Exception {
        int indexHeading = styler.getIndex(sOriginal.getBasedOn().getVal());
        String heading = styler.getHeadingProperty(indexHeading, "name");

        if (Helper.isHeading(heading, sResponse.getBasedOn().getVal(), sResponse.getStyleId())) {

            String fontname, size, bold, hexcolor, line_spacing, spacing_before, spacing_after;
            String queryd, querys, queryb, query1, query2, query3;
            int values, check3, check4, check5, check6, check7 = 0, check8 = 0, check9 = 0;

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
//            System.out.print("\t\tFont Name:");
            check3 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:sz[@w:val and string-length(@w:val)!=0]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:sz[contains(@w:val," + size + ")]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:sz[contains(@w:val," + size + ")]]]";
//            System.out.print("\t\tFont Size:");
            check4 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:b[not(@*) or w:val='true']]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:b[not(@*) or w:val='true']]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:b[not(@*) or w:val='true']]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:b[not(@*)]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:b[not(@*)]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:b[not(@*)]]]";
//            System.out.print("\t\tBold:");
            check5 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:color[@w:val]]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:color[@w:val]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:color[@w:val]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:rPr[w:color[ contains(@w:val,'" + hexcolor + "') ]]]";
//            System.out.print("\t\tColor:");
            check6 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:before]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[@w:before]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[@w:before]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[contains(@w:before,'" + spacing_before + "')]]]";
//            System.out.print("\t\tSpacing Before:");
            check7 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            queryd = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[@w:after]]]";
            querys = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[@w:after]]]";
            queryb = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[@w:after]]]";

            query1 = "//w:p[@w14:paraId='" + pResponse.getParaId() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
            query2 = "//w:style[@w:styleId='" + sResponse.getStyleId() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
            query3 = "//w:style[@w:styleId='" + sResponse.getBasedOn().getVal() + "' and w:pPr[w:spacing[contains(@w:after,'" + spacing_after + "')]]]";
//            System.out.print("\t\tSpacing After:");
            check8 = matchStyle(query1, query2, query3, queryd, querys, queryb);

            return (double) (check3 + check4 + check5 + check6 + check7 + check8) / values >= Verifier.TOC_LIMIT_STYLE;
        }

        return false;
    }

    private boolean checkListing(Style sOriginal, Style sResponse, P pResponse) throws Exception {
        String query1 = "//w:p[contains(@w14:paraId,'" + pResponse.getParaId() + "') and w:pPr[w:numPr]]";
        String query2 = "//w:style[contains(@w:styleId,'" + sResponse.getStyleId() + "') and w:pPr[w:numPr]]";
        String query3 = "//w:style[contains(@w:styleId,'" + sResponse.getBasedOn().getVal() + "') and w:pPr[w:numPr]]";
//        System.out.print("\t\tListing:");
        return matchStyle(query1, query2, query3, query1, query2, query3) == 1;
    }

    public void validateTOC() throws Exception {

        LinkedList tocOriginal = loadTOC(Verifier.INDEX_ORIGINAL);
        LinkedList tocResponse = loadTOC(Verifier.INDEX_RESPONSE);

        LinkedList headingsOriginal = loadHeadings(Verifier.INDEX_ORIGINAL);
        LinkedList headingsResponse = loadHeadings(Verifier.INDEX_RESPONSE);

        System.out.println("Grading: Table of Contents");

        int grade = 0;

        //TOC: exist or not
        if (tocResponse.size() > 0) {
            grade = 5;
            System.out.println("\tHas TOC +5");

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

            missing = tocOriginal.size() - sameInOriginal;
            notHere = tocResponse.size() - sameInOriginal;
            totalTOC = sameInOriginal - notHere - missing;

            if ((double) totalTOC / tocOriginal.size() >= Verifier.TOC_LIMIT_ELEMENTSINTOC) {
                grade += 10;
                System.out.println("\tMost elements in TOC! +10");
            } else {
                grade += 3;
                System.out.println("\tFew elements in TOC +3");
            }

        } else {
            System.out.println("\tWithout TOC +0");
        }

        //Styles: Same headingsOriginal on sResponse file
        P pResponse;
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
                    pResponse = (org.docx4j.wml.P) o2;
                    sOriginal = getStyleByHeadingName(Verifier.INDEX_ORIGINAL, o);
                    sResponse = getStyleByHeadingName(Verifier.INDEX_RESPONSE, o2);
                    sameStyle = checkStyle(sOriginal, sResponse, pResponse);
                    hasListing = checkListing(sOriginal, sResponse, pResponse);
//                    System.out.println("\t"+Helper.getTextFromP(pResponse.getContent())+" "+pResponse.getPPr().getPStyle().getVal());
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

        if ((double) foundStyle / headingsOriginal.size() >= Verifier.TOC_LIMIT_HEADINGS) {
            grade += 5;
            System.out.println("\tMost with correct styles! +5");
        } else {
            System.out.println("\tFew headings with correct styles +0");
        }

        //Listing
        if ((double) foundListing / headingsOriginal.size() >= Verifier.TOC_LIMIT_HEADINGS) {
            grade += 5;
            System.out.println("\tMost with correct listing! +5");
        } else {
            System.out.println("\tFew headings with correct listing +0");
        }

        System.out.println("\tGrade: " + grade + "/" + Verifier.GRADE_TOC);
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

        String query;

        query = "//w:p[w:r[w:rPr[w:rStyle[@w:val='FootnoteReference']]]]";

        LinkedList footNoteOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList footNoteResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        HashMap<String, String> hashOriginal = getFootNotes(Verifier.INDEX_ORIGINAL, footNoteOriginal);
        HashMap<String, String> hashResponse = getFootNotes(Verifier.INDEX_RESPONSE, footNoteResponse);

        int countKeyOriginal = hashOriginal.keySet().size();
        int countKeyResponse = hashResponse.keySet().size();

        System.out.println("Grading: Footnotes");

        int grade = 0;
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

        if ((double) count_specs / total_specs > Verifier.FOOTNOTE_LIMIT_2) {
            grade = 7;
            System.out.println("\tMost Specs! +" + grade);
        } else if ((double) count_specs / total_specs >= Verifier.FOOTNOTE_LIMIT_1) {
            grade = 4;
            System.out.println("\t40-89% Specs! +" + grade);
        } else if ((double) count_specs / total_specs > 0 && (double) count_specs / total_specs < Verifier.FOOTNOTE_LIMIT_1) {
            grade = 1;
            System.out.println("\t 0-39% Specs! +" + grade);
        } else {
            grade = 0;
            System.out.println("\tNone! +" + grade);
        }

        System.out.println("\tGrade: " + grade + "/" + Verifier.GRADE_FOOTNOTE);
    }

    private void validateDropCap() throws Exception {
        String query = "//w:p[w:pPr[w:framePr[@w:dropCap]]] | //w:p[w:pPr[w:framePr[@w:dropCap]]]/following-sibling::*[1]";
        LinkedList dropCapOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList dropCapResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        System.out.println("Grading: Capital Letters");

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
            txtR = Helper.getTextFromP(((P) rOv).getContent()).substring(0, 20).trim();

//            System.out.println(rOp.getPPr().getFramePr().getDropCap().value()+" "+ capR+" "+txtR);
            for (Iterator it2 = dropCapOriginal.iterator(); it2.hasNext();) {
                pOp = (P) it2.next();
                capO = Helper.getTextFromP(((P) pOp).getContent()).trim();
                pOv = (P) it2.next();
                txtO = Helper.getTextFromP(((P) pOv).getContent()).substring(0, 20).trim();

                //match on text
                if (capR.compareTo(capO) == 0 && txtR.compareTo(txtO) == 0) {

//                    System.out.println("\t"+pOp.getPPr().getFramePr().getDropCap().value()+" "+capO+" "+txtO);
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

//        System.out.println((double) counterDL / capCounterOriginal.doubleValue() + " " + (double) counterD / capCounterOriginal.doubleValue() + " " + (double) counterML / capCounterOriginal.doubleValue() + " " + (double) counterM / capCounterOriginal.doubleValue());
        int grade = 0;

        if ((double) counterDL / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_ALL) {
            grade += 6;
            System.out.println("\tMost Specs! +" + grade);
        } else if ((double) counterD / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_SOME) {
            grade += 5;
            System.out.println("\t66% - 89%! +" + grade);
        } else if ((double) counterD / capCounterOriginal.doubleValue() < Verifier.CD_LIMIT_SOME && (double) counterML / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_SOME && (double) counterML / capCounterOriginal.doubleValue() <= Verifier.CD_LIMIT_ALL) {
            grade += 4;
            System.out.println("\t40% - 65%! +" + grade);
        } else if ((double) counterML / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_ALL && (double) counterD / capCounterOriginal.doubleValue() < Verifier.CD_LIMIT_SOME) {
            grade += 2;
            System.out.println("\t11% - 39%! +" + grade);
        } else if ((double) counterM / capCounterOriginal.doubleValue() >= Verifier.CD_LIMIT_SOME && (double) counterM / capCounterOriginal.doubleValue() <= Verifier.CD_LIMIT_ALL) {
            grade += 1;
            System.out.println("\t 0% - 10%! +" + grade);
        } else {
            grade += 0;
        }

        System.out.println("\tGrade: " + grade + "/" + Verifier.GRADE_CAP);
    }

    private void validateColumns() throws Exception {
        int specs = 0, totalSpecs = 0;
        String nquery, textO, textR;
        int countDifferentText;
        P pnO, pnR, pO1, pO2, pR1, pR2;
        int grade = 0;

        String query = "//w:p[w:pPr[w:sectPr[w:cols[@w:num]]]] | //w:p[w:pPr[w:sectPr[w:cols[@w:num]]]]/preceding-sibling::w:p[w:pPr[w:sectPr[w:cols]]][1]";
        LinkedList dropCapOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList dropCapResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        System.out.println("Grading: Columns");

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

//                System.out.println((double) similars.doubleValue() / linsideR.size());
                //Similar columns/total > threshold
                //If and only if they're similars paragraphs then check others specs
                if ((double) similars.doubleValue() / linsideR.size() >= Verifier.COLUMN_MINIMUM_SAME_PARAGRAPHS) {

                    found = true;

                    //Same number of columns by columns group
                    totalSpecs++;
                    specs = (pO2.getPPr().getSectPr().getCols().getNum().intValue() == pR2.getPPr().getSectPr().getCols().getNum().intValue()) ? specs + 1 : specs - 1;
//                    System.out.println((pO2.getPPr().getSectPr().getCols().getNum().intValue() == pR2.getPPr().getSectPr().getCols().getNum().intValue()) );

                    //Same number or paragraphs inside
                    totalSpecs++;
                    specs = (linsideO.size() == linsideR.size()) ? specs + 1 : specs - 1;
//                    System.out.println(linsideO.size() == linsideR.size());

                    //Same separator by groups
                    totalSpecs++;
                    specs = (XmlUtils.marshaltoString(pO2, true, true).contains("w:sep") && XmlUtils.marshaltoString(pR2, true, true).contains("w:sep")) ? specs + 1 : specs - 1;
//                    System.out.println(XmlUtils.marshaltoString(pO2, true, true).contains("w:sep") && XmlUtils.marshaltoString(pR2, true, true).contains("w:sep"));

                    totalSpecs++;
                    specs = (similars.intValue() == linsideR.size()) ? specs + 1 : specs - 1;
//                    System.out.println("Same Texts: "+similars.intValue() +" "+ linsideR.size());
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

//        System.out.println(specs+" "+totalSpecs);
        if (specs < 0) {
            grade = 0;
            System.out.println("\tNone! " + grade);
        } else {
            if ((double) specs / totalSpecs >= Verifier.COLUMN_LIMIT_2) {
                grade += 10;
                System.out.println("\tMost Specs! +" + grade);
            } else if ((double) specs / totalSpecs >= Verifier.COLUMN_LIMIT_1) {
                grade += 5;
                System.out.println("\t40% - 89%! +" + grade);
            } else {
                grade += 1;
                System.out.println("\t 0% - 39%! +" + grade);
            }
        }

        System.out.println("\tGrade: " + grade + "/" + Verifier.GRADE_COLUMNS);

    }

    private void validateBdr() throws Exception {
        String query = "//w:p[w:pPr[w:pBdr]]";
        LinkedList txtBdrOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList txtBdrResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        System.out.println("Grading: Borders");
        int grade = 0;

        //Same number of borderO paragraphs
        double spec_numberPO = 0, spec_numberPR = 0;
        spec_numberPO = 1;
        if (txtBdrOriginal.size() == txtBdrResponse.size()) {
            spec_numberPR = 1;
        }
//        System.out.println(spec_numberPR + " " + spec_numberPO);

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
                if (txtO.trim().compareTo(txtR.trim()) == 0 || (i.intValue() == j && i.intValue() == Verifier.NAME_POSITION_BDR && txtR.length() > 0)) {
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

//                            System.out.println("Top");
//                            System.out.println("O: " + borderO.getVal().value());
//                            System.out.println("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                System.out.println("T-Same " + colorO + " " + colorR);
                            } else {
//                                System.out.println("T-Diff " + colorO + " " + colorR);
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

//                            System.out.println("Bottom");
//                            System.out.println("O: " + borderO.getVal().value());
//                            System.out.println("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                System.out.println("B-Same " + colorO + " " + colorR);
                            } else {
//                                System.out.println("B-Diff " + colorO + " " + colorR);
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

//                            System.out.println("Left");
//                            System.out.println("O: " + borderO.getVal().value());
//                            System.out.println("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                System.out.println("L-Same " + colorO + " " + colorR);
                            } else {
//                                System.out.println("L-Diff " + colorO + " " + colorR);
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

//                            System.out.println("Right");
//                            System.out.println("O: " + borderO.getVal().value());
//                            System.out.println("R: " + borderR.getVal().value());
                            if (borderO.getVal().value().compareTo(borderR.getVal().value()) == 0) {
                                countSameBTypeR.getAndIncrement();
                            }

                            colorO = borderO.getColor().compareTo("auto") == 0 ? "000000" : borderO.getColor();
                            colorR = borderR.getColor().compareTo("auto") == 0 ? "000000" : borderR.getColor();

                            if (Helper.similarTo(colorO, colorR, 1.0)) {
                                countSameColorR.getAndIncrement();
//                                System.out.println("R-Same " + colorO + " " + colorR);
                            } else {
//                                System.out.println("R-Diff " + colorO + " " + colorR);
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
//        System.out.println("Equal text: " + countEqualTxtR + " " + countEqualTxtO + " " + spec_samePR);

        //Same border in paragraphs
        double spec_sameBO = 0, spec_sameBR = 0;
        spec_sameBO = 1;
        spec_sameBR = (double) (countSameBTypeR.doubleValue() / (countSameBTypeO.doubleValue() == 0 ? 1 : countSameBTypeO.doubleValue()));
//        System.out.println("Equal border: " + countSameBTypeR + " " + countSameBTypeO + " " + spec_sameBR);

        //Same border color in paragraphs
        double spec_sameCO = 0, spec_sameCR = 0;
        spec_sameCO = 1;
        spec_sameCR = (double) (countSameColorR.doubleValue() / (countSameColorO.doubleValue() == 0 ? 1 : countSameColorO.doubleValue()));
//        System.out.println("Equal color: " + countSameColorR + " " + countSameColorO + " " + spec_sameCR);

        //Same shading in paragraphs
        double spec_sameSHO = 0, spec_sameSHR = 0;
        spec_sameSHO = 1;
        spec_sameSHR = (double) (countSameShadingR.doubleValue() / (countSameShadingO.doubleValue() == 0 ? 1 : countSameShadingO.doubleValue()));
//        System.out.println("Equal shading: " + countSameShadingR + " " + countSameShadingO + " " + spec_sameSHR);

        //Images
        query = "//w:p[.//pic:spPr[.//a:prstDash or .//a:srgbClr]]//pic:pic";
        LinkedList imagesOriginal = getDocumentObjectByQuery(Verifier.INDEX_ORIGINAL, query);
        LinkedList imagesResponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        double spec_numberIO = 0, spec_numberIR = 0;
        spec_numberIO = 1;
        if (imagesOriginal.size() == imagesResponse.size()) {
            spec_numberIR = 1;
        }
//        System.out.println(spec_numberIR + " " + spec_numberIO);

        AtomicInteger sameNameImgO = new AtomicInteger(0), sameNameImgR = new AtomicInteger(0);
        AtomicInteger sameDashImgO = new AtomicInteger(0), sameDashImgR = new AtomicInteger(0);

        imagesOriginal.stream().forEach((objO) -> {
//        for (int i = 0; i < imagesOriginal.size(); i++) {
            Pic picO = (Pic) ((javax.xml.bind.JAXBElement) objO).getValue(); //imagesOriginal.get(i)).getValue();
            String picNameO = picO.getNvPicPr().getCNvPr().getDescr();

            sameNameImgO.getAndIncrement();
            for (int j = 0; j < imagesResponse.size(); j++) {
                Pic picR = (Pic) ((javax.xml.bind.JAXBElement) imagesResponse.get(j)).getValue();
                String picNameR = picR.getNvPicPr().getCNvPr().getDescr();

                if (picNameO.compareTo(picNameR) == 0) {
                    sameNameImgR.getAndIncrement();

                    String dashO = picO.getSpPr().getLn().getPrstDash().getVal().value();
                    String dashR = picR.getSpPr().getLn().getPrstDash().getVal().value();

                    sameDashImgO.getAndIncrement();
                    if (dashO.compareTo(dashR) == 0) {
                        sameDashImgR.getAndIncrement();
                    }

                }
            }
        });

        double spec_sameIO = 0, spec_sameIR = 0;

        spec_sameIO = 1;
        spec_sameIR = (double) (sameNameImgR.doubleValue() / (sameNameImgO.doubleValue() == 0 ? 1 : sameNameImgO.doubleValue()));
//        System.out.println("Equal image: " + sameNameImgR + " " + sameNameImgO + " " + spec_sameIR);

        double spec_sameDashO = 0, spec_sameDashR = 0;

        spec_sameDashO = 1;
        spec_sameDashR = (double) (sameDashImgR.doubleValue() / (sameDashImgO.doubleValue() == 0 ? 1 : sameDashImgO.doubleValue()));
//        System.out.println("Equal dash: " + sameDashImgR + " " + sameDashImgO + " " + spec_sameDashR);

        double totalO = spec_numberIO + spec_numberPO + spec_samePO + spec_sameIO + ((spec_sameDashO) / (spec_sameIO == 0 ? 1 : spec_sameIO)) + ((spec_sameBO + spec_sameCO + spec_sameSHO) / (spec_samePO == 0 ? 1 : spec_samePO));
        double totalR = spec_numberIR + spec_numberPR + spec_samePR + spec_sameIR + ((spec_sameDashR) / (spec_sameIR == 0 ? 1 : spec_sameIR)) + ((spec_sameBR + spec_sameCR + spec_sameSHR) / (spec_samePR == 0 ? 1 : spec_samePR));

//        System.out.println(totalR + " " + totalO);
        if ((double) totalR / totalO > Verifier.BDR_LIMIT_1) {
            grade += 10;
            System.out.println("\tMost Specs! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_2) {
            grade += 9;
            System.out.println("\t66% - 89%! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_3) {
            grade += 7;
            System.out.println("\t40% - 65%! +" + grade);
        } else if ((double) totalR / totalO > Verifier.BDR_LIMIT_4) {
            grade += 4;
            System.out.println("\t11% - 39%! +" + grade);
        } else {
            grade += 1;
            System.out.println("\t 0% - 10%! +" + grade);
        }

        System.out.println("\tGrade: " + grade + "/" + Verifier.GRADE_BORDER);

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

    public void validateFooter() throws Exception {

        int specs = 0, totalSpecs = 0;
        double grade = 0;
        boolean hasDefault = false, hasEven = false, hasFirst = false;

        FooterPart elOriginal = null, elResponse = null;
        HeaderFooterPolicy hfp;

        Iterator<FooterPart> originalFooter, responseFooter;

        int type = Verifier.FOOTER_DEFAULT;

        originalFooter = getFooter(Verifier.INDEX_ORIGINAL, type).iterator();

        while (originalFooter.hasNext()) {
            elOriginal = originalFooter.next();

            if (elOriginal != null) {

                //Verify numbered footer
                totalSpecs = isFooterNumbered(elOriginal) ? totalSpecs + 1 : totalSpecs;
                totalSpecs = getContentFooter(elOriginal) != null ? totalSpecs + 1 : totalSpecs;
                System.out.println("O numbered: " + isFooterNumbered(elOriginal));
                System.out.println("O content: " + getContentFooter(elOriginal));

            }

            responseFooter = getFooter(Verifier.INDEX_RESPONSE, type).iterator();
            while (responseFooter.hasNext()) {
                elResponse = responseFooter.next();

                if (elResponse != null) {
                    specs = isFooterNumbered(elResponse) ? (isFooterNumbered(elOriginal) == isFooterNumbered(elResponse) ? specs + 1 : specs - 1) : specs;
                    specs = getContentFooter(elResponse) != null ? (getContentFooter(elOriginal).equals(getContentFooter(elResponse)) ? specs + 1 : specs - 1) : specs;
                    System.out.println("R numbered: " + isFooterNumbered(elResponse));
                    System.out.println("R content: " + getContentFooter(elResponse));
                }
            }
        }

        //Has footer
        totalSpecs++;
        if (hasDefault | hasEven | hasFirst) {
            specs++;
        }

        System.out.println(totalSpecs + " :: " + specs);

    }

    public void validate() throws Exception {
        loadDocument(Verifier.INDEX_ORIGINAL);
        loadDocument(Verifier.INDEX_RESPONSE);

//        validateTOC();
//        validateBdr();
//        validateFootNote();
//        validateDropCap();
//        validateColumns();
//        validateFooter();
//        showXML();
        show();
//        showFooterXML();
//        pagecount();
    }

    private void show() throws Exception {
        String query = "//w:sectPr[w:footerReference]/w:footerReference";
        LinkedList lresponse = getDocumentObjectByQuery(Verifier.INDEX_RESPONSE, query);

        FooterReference r;

        for (Object o1 : lresponse) {
            r = (org.docx4j.wml.FooterReference) o1;
            System.out.println(r.getId() + " " + r.getType().value());

            //
            RelationshipsPart rp = wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getRelationshipsPart();

            List<Relationship> rels = rp.getRelationshipsByType(Namespaces.FOOTER);
            Iterator<Relationship> it = rels.iterator();

            while (it.hasNext()) {
                Relationship rel = it.next();
                JaxbXmlPart part = (JaxbXmlPart) rp.getPart(rel);
                if (rel.getId().equals(r.getId())) {
                    System.out.println(Helper.getTextFromFtr(part.getContents()));
                }
            }
            //
        }
    }

    private void showFooterXML() throws Exception {

        //Relaciones intermedias con los otros archivos
        RelationshipsPart rp = wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getRelationshipsPart();

//        System.out.println(rp.getXML());
        List<Relationship> rels = rp.getRelationshipsByType(Namespaces.FOOTER);
        Iterator<Relationship> it = rels.iterator();
        while (it.hasNext()) {
            Relationship rel = it.next();
            JaxbXmlPart part = (JaxbXmlPart) rp.getPart(rel);
            System.out.println(rel.getId());
            System.out.println(part.getContents());
        }

    }

    private void showXML() throws Exception {
        //.getStyleDefinitionsPart()
        System.out.println(wordMLPackage[Verifier.INDEX_RESPONSE].getMainDocumentPart().getXML());
//        
//        String query = "//w:styleBody";
//        LinkedList body = getStyleObjectByQuery(Verifier.INDEX_ORIGINAL, query);
//        body.stream().forEach((obj) -> {
//            System.out.println(XmlUtils.marshaltoString(obj, true, true));
//        });
    }

}

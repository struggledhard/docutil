package com.gzcdc.officeserver.util;

import java.io.*;
import org.apache.poi.xwpf.usermodel.*; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr; 
import java.util.List; 
import java.util.Iterator; 
import java.util.Stack; 
import org.apache.xmlbeans.XmlCursor; 
import org.apache.xmlbeans.XmlException; 
import org.w3c.dom.Node; 
import org.w3c.dom.NodeList; 

/** 
 * Second attempt at inserting text at a bookmark defined within a Word 
 * document. Note that there is one SERIOUS limitations with the code as it 
 * stands; at least only one as far as I am aware: nested bookmarks. 
 * 
 * It is possible to create a document and to nest one bookmark within another. 
 * Typically, a bookmark is inserted into a piece of text, that is then selected 
 * and another bookmark is added to that selection. The XML markup might look 
 * something like this 
 * 
 * <pre>
 * <w:p w:rsidR="00945150" w:rsidRDefault="00945150">
 *   <w:r>
 *     <w:t xml:space="preserve">
 *     Imagine I want to insert one bookmark at the start of this 
 *     </w:t>
 *   </w:r>
 *     <w:bookmarkStart w:id="0" w:name="OUTER"/>
 *       <w:r>
 *         <w:t xml:space="preserve">piece of text and another just 
 *         </w:t>
 *     </w:r>
 *   <w:proofErr w:type="gramStart"/>
 *   <w:r>
 *     <w:t xml:space="preserve">here 
 *     </w:t>
 *   </w:r>
 *   <w:bookmarkStart w:id="1" w:name="INNER"/>
 *   <w:bookmarkEnd w:id="1"/>
 *     <w:r>
 *       <w:t>. 
 *       </w:t>
 *     </w:r>
 *   <w:bookmarkEnd w:id="0"/>
 *   <w:proofErr w:type="gramEnd"/>
 * </w:p>
 * </pre>
 * 
 * Using Word macros to conduct tests, a number of things have become apparent. 
 * The first is that within the Word Object Model, a Bookmark is considered to 
 * be a Range object and this limits the operations that can be performed on it. 
 * It is quite possible to insert some text in front of the bookmarks start tag 
 * or behind the bookmarks end tag. The text is never inserted into the markup 
 * between the bookmarkStart and bookmarkEnd rags and, therefore, will not 
 * replace any text that does already appear between them. With regard to 
 * styling, it seems as though the following guidelines hold true; 
 * 
 * 1. If text is being inserted before the bookmark, then it will 'inherit' any 
 * style information from the character run that immediately precedes it, if 
 * any. 2. If the text is being inserted after the bookmarkEnd tag then it will 
 * inherit it's styling from the nearest run element that is contained between 
 * the bookmarkEnd tag and it's matching bookmarkStart tag, if any. 
 * 
 * Currently, I am unsure on a couple of points; 
 * 
 * 1. Whether it is possible for a document to contain two or more bookmarks 
 * with the same name. Initial testing suggested that this is not possible but 
 * the code has been written as if it is. That is to say that once a bookmark is 
 * found the search for a subsequent bookmark with the same name will continue. 
 * This behaviour is easy to amend however. 
 * 2. Should the code offer a third 
 * option, to replace the text, if any, contained between the bookmarkStart and 
 * bookmarkEnd tags? If so, what should happen to any bookmarks that are 
 * contained between the start and end tags? 
 * 
 * @author Mark Beardsley 
 * @version 1.00 16th June 2012 
 *          1.10 20th June 2012 - Added the ability to replace the text between the opening 
 *                                and closing brackets ([ and ]) that appear when the document 
 *                                is open in Word. 
 */ 
public class DOCXTest { 

    public static final int INSERT_BEFORE = 0; 
    public static final int INSERT_AFTER = 1; 
    public static final int REPLACE = 2; 
    private XWPFDocument document = null; 

    public DOCXTest() { 
    } 

    /** 
     * Opens a Word OOXML file. 
     * 
     * @param filename An instance of the String class that encapsulates the 
     * path to and name of a Word OOXML (.docx) file. 
     * @throws IOException Thrown if a problem occurs within the underlying file 
     * system. 
     */ 
    public final void openFile(String filename) throws IOException { 
        File file = null; 
        FileInputStream fis = null; 
        try { 
            // Simply open the file and store a reference into the 'document' 
            // local variable. 
            file = new File(filename); 
            fis = new FileInputStream(file); 
            this.document = new XWPFDocument(fis); 
        } finally { 
            try { 
                if (fis != null) { 
                    fis.close(); 
                    fis = null; 
                } 
            } catch (IOException ioEx) { 
                // Swallow this exception. It would have occured onyl 
                // when releasing the file handle and should not pose 
                // problems to later processing. 
            } 
        } 
    } 

    /** 
     * Saves a Word OOXML file away under the name, and to the location, 
     * specified. 
     * 
     * @param filename An instance of the String class that encapsulates the of 
     * the file and the location into which it should be stored. 
     * @throws IOException Thrown if a problem occurs in the underlying file 
     * system. 
     */ 
    public final void saveAs(String filename) throws IOException { 
        File file = null; 
        FileOutputStream fos = null; 
        try { 
            file = new File(filename); 
            fos = new FileOutputStream(file); 
            this.document.write(fos); 
        } finally { 
            if (fos != null) { 
                fos.close(); 
                fos = null; 
            } 
        } 
    } 

    /** 
     * Inserts a value at a location within the Word document specified by a 
     * named bookmark. 
     * 
     * @param bookmarkName An instance of the String class that encapsulates the 
     * name of the bookmark. Note that case is important and the case of the 
     * bookmarks name within the document and that of the value passed to this 
     * parameter must match. 
     * @param bookmarkValue An instance of the String class that encapsulates 
     * the value that should be inserted into the document at the location 
     * specified by the bookmark. 
     * @param where A primitive int whose value indicates whether the text 
     * should be inserted before or after the bookmark. Note that constants have 
     * been defined - DOCXTest.INSERT_BEFORE and DOCXTest.INSERT_AFTER - for 
     * this purpose. 
     */ 
    public final void insertAtBookmark(String bookmarkName, 
            String bookmarkValue, int where) throws XmlException { 
        List<XWPFTable> tableList = null; 
        Iterator<XWPFTable> tableIter = null; 
        List<XWPFTableRow> rowList = null; 
        Iterator<XWPFTableRow> rowIter = null; 
        List<XWPFTableCell> cellList = null; 
        Iterator<XWPFTableCell> cellIter = null; 
        XWPFTable table = null; 
        XWPFTableRow row = null; 
        XWPFTableCell cell = null; 

        // Firstly, deal with any paragraphs in the body of the document. 
        this.procParaList(this.document.getParagraphs(), bookmarkName, bookmarkValue, where); 

        // Then check to see if there are any bookmarks in table cells. To do this 
        // it is necessary to get at the list of paragraphs 'stored' within the 
        // individual table cell, hence this code which get the tables from the 
        // document, the rows from each table, the cells from each row and the 
        // paragraphs from each cell. 
        tableList = this.document.getTables(); 
        tableIter = tableList.iterator(); 
        while (tableIter.hasNext()) { 
            table = tableIter.next(); 
            rowList = table.getRows(); 
            rowIter = rowList.iterator(); 
            while (rowIter.hasNext()) { 
                row = rowIter.next(); 
                cellList = row.getTableCells(); 
                cellIter = cellList.iterator(); 
                while (cellIter.hasNext()) { 
                    cell = cellIter.next(); 
                    this.procParaList(cell.getParagraphs(), 
                            bookmarkName, 
                            bookmarkValue, 
                            where); 
                } 
            } 
        } 
    } 

    /** 
     * Inserts text into the document at the position indicated by a specific 
     * bookmark. Note that the current implementation does not take account of 
     * nested bookmarks, that is bookmarks that contain other bookmarks. Note 
     * also that any text contained within the bookmark itself will be removed. 
     * 
     * @param paraList An instance of a class that implements the List interface 
     * and which encapsulates references to one or more instances of the 
     * XWPFParagraph class. 
     * @param bookmarkName An instance of the String class that encapsulates the 
     * name of the bookmark that identifies the position within the document 
     * some text should be inserted. 
     * @param bookmarkValue An instance of the AString class that encapsulates 
     * the text that should be inserted at the location specified by the 
     * bookmark. 
     * @param where A primitive int whose value indicates where the text should 
     * be inserted relative to the bookmark, i.e. before or after the bookmark. 
     */ 
    private final void procParaList(List<XWPFParagraph> paraList, 
            String bookmarkName, String bookmarkValue, int where) throws XmlException { 
        Iterator<XWPFParagraph> paraIter = null; 
        XWPFParagraph para = null; 
        List<CTBookmark> bookmarkList = null; 
        Iterator<CTBookmark> bookmarkIter = null; 
        CTBookmark bookmark = null; 
        XWPFRun run = null; 

        // Get an Iterator for the XWPFParagraph object and step through them 
        // one at a time. 
        paraIter = paraList.iterator(); 
        while (paraIter.hasNext()) { 
            para = paraIter.next(); 

            // Get a List of the CTBookmark object sthat the paragraph 
            // 'contains' and step through these one at a time. 
            bookmarkList = para.getCTP().getBookmarkStartList(); 
            bookmarkIter = bookmarkList.iterator(); 
            while (bookmarkIter.hasNext()) { 
                bookmark = bookmarkIter.next(); 

                // If the name of the CTBookmakr object matches the value 
                // encapsulated within the argumnet passed to the bookmarkName 
                // parameter then this is where the text should be inserted. 
                if (bookmark.getName().equals(bookmarkName)) { 

                    // Create a new character run to hold the value encapsulated 
                    // within the argument passed to the bookmarkValue parameter 
                    // and then test whether this new run shouold be inserted 
                    // into the document before or after the bookmark. 
                    run = para.createRun(); 
                    run.setText(bookmarkValue); 
                    switch (where) { 
                        case DOCXTest.INSERT_AFTER: 
                            this.insertAfterBookmark(bookmark, run, para); 
                            break; 
                        case DOCXTest.INSERT_BEFORE: 
                            this.insertBeforeBookmark(bookmark, run, para); 
                            break; 
                        case DOCXTest.REPLACE: 
                            this.replaceBookmark(bookmark, run, para); 
                            break; 

                    } 
                } 
            } 
        } 
    } 

    /** 
     * Inserts some text into a Word document in a position that is immediately 
     * after a named bookmark. 
     * 
     * Bookmarks can take two forms, they can either simply mark a location 
     * within a document or they can do this but contain some text. The 
     * difference is obvious from looking at some XML markup. The simple 
     * placeholder bookmark will look like this; 
     * 
     * <pre>
     * 
     * <w:bookmarkStart w:name="AllAlone" w:id="0"/><w:bookmarkEnd w:id="0"/>
     * 
     * </pre>
     * 
     * Simply a pair of tags where one tag has the name bookmarkStart, the other 
     * the name bookmarkEnd and both share matching id attributes. In this case, 
     * the text will simply be inserted into the document at a point immediately 
     * after the bookmarkEnd tag. No styling will be applied to the text, it 
     * will simply inherit the documents defaults. 
     * 
     * The more complex case looks like this; 
     * 
     * <pre>
     * 
     * <w:bookmarkStart w:name="InStyledText" w:id="3"/>
     *   <w:r w:rsidRPr="00DA438C">
     *     <w:rPr>
     *       <w:rFonts w:hAnsi="Engravers MT" w:ascii="Engravers MT" w:cs="Arimo"/>
     *       <w:color w:val="FF0000"/>
     *     </w:rPr>
     *     <w:t>text</w:t>
     *   </w:r>
     * <w:bookmarkEnd w:id="3"/>
     * 
     * </pre>
     * 
     * Here, the user has selected the word 'text' and chosen to insert a 
     * bookmark into the document at that point. So, the bookmark tags 'contain' 
     * a character run that is styled. Inserting any text after this bookmark, 
     * it is important to ensure that the styling is preserved and copied over 
     * to the newly inserted text. 
     * 
     * The approach taken to dealing with both cases is similar but slightly 
     * different. In both cases, the code simply steps along the document nodes 
     * until it finds the bookmarkEnd tag whose ID matches that of the 
     * bookmarkStart tag. Then, it will look to see if there is one further node 
     * following the bookmarkEnd tag. If there is, it will insert the text into 
     * the paragraph immediately in front of this node. If, on the other hand, 
     * there are no more nodes following the bookmarkEnd tag, then the new run 
     * will simply be positioned at the end of the paragraph. 
     * 
     * Styles are dealt with by 'looking' for a 'w:rPr' element whilst iterating 
     * through the nodes. If one is found, its details will be captured and 
     * applied to the run before the run is inserted into the paragraph. If 
     * there are multiple runs between the bookmarkStart and bookmarkEnd tags 
     * and these have different styles applied to them, then the style applied 
     * to the last run before the bookmarkEnd tag - if any - will be cloned and 
     * applied to the newly inserted text. 
     * 
     * @param bookmark An instance of the CTBookmark class that encapsulates 
     * information about the bookmark. 
     * @param run An instance of the XWPFRun class that encapsulates the text 
     * that is to be inserted into the document following the bookmark. 
     * @param para An instance of the XWPFParagraph class that encapsulates that 
     * part of the document, a paragraph, into which the run will be inserted. 
     */ 
    private void insertAfterBookmark(CTBookmark bookmark, XWPFRun run, 
            XWPFParagraph para) { 
        Node nextNode = null; 
        Node insertBeforeNode = null; 
        Node styleNode = null; 
        int bookmarkStartID = 0; 
        int bookmarkEndID = -1; 

        // Capture the id of the bookmarkStart tag. The code will step through 
        // the document nodes 'contained' within the start and end tags that have 
        // matching id numbers. 
        bookmarkStartID = bookmark.getId().intValue(); 

        // Get the node for the bookmark start tag and then enter a loop that 
        // will step from one node to the next until the bookmarkEnd tag with 
        // a matching id is fouind. 
        nextNode = bookmark.getDomNode(); 
        while (bookmarkStartID != bookmarkEndID) { 

            // Get the next node along and check to see if it is a bookmarkEnd 
            // tag. If it is, get its id so that the containing while loop can 
            // be terminated once the correct end tag is found. Note that the 
            // id will be obtained as a String and must be converted into an 
            // integer. This has been coded to fail safely so that if an error 
            // is encuntered converting the id to an int value, the while loop 
            // will still terminate. 
            nextNode = nextNode.getNextSibling(); 
            if (nextNode.getNodeName().contains("bookmarkEnd")) { 
                try { 
                    bookmarkEndID = Integer.parseInt( 
                            nextNode.getAttributes().getNamedItem("w:id").getNodeValue()); 
                } catch (NumberFormatException nfe) { 
                    bookmarkEndID = bookmarkStartID; 
                } 
            } // If we are not dealing with a bookmarkEnd node, are we dealing 
            // with a run node that MAY contains styling information. If so, 
            // then get that style information from the run. 
            else { 
                if (nextNode.getNodeName().equals("w:r")) { 
                    styleNode = this.getStyleNode(nextNode); 
                } 
            } 
        } 

        // After the while loop completes, it should have located the correct 
        // bookmarkEnd tag but we cannot perform an insert after only an insert 
        // before operation and must, therefore, get the next node. 
        insertBeforeNode = nextNode.getNextSibling(); 

        // Style the newly inserted text. Note that the code copies or clones 
        // the style it found in another run, failure to do this would remove the 
        // style from one node and apply it to another. 
        if (styleNode != null) { 
            run.getCTR().getDomNode().insertBefore( 
                    styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild()); 
        } 

        // Finally, check to see if there was a node after the bookmarkEnd 
        // tag. If there was, then this code will insert the run in front of 
        // that tag. If there was no node following the bookmarkEnd tag then the 
        // run will be inserted at the end of the paragarph and this was taken 
        // care of at the point of creation. 
        if (insertBeforeNode != null) { 
            para.getCTP().getDomNode().insertBefore( 
                    run.getCTR().getDomNode(), insertBeforeNode); 
        } 
    } 

    /** 
     * Inserts some text into a Word document immediately in front of the 
     * location of a named bookmark. 
     * 
     * This case is slightly more straightforward than inserting after the 
     * bookmark. For example, it is possible only to insert a new node in front 
     * of an existing node. When inserting after the bookmark, then end node had 
     * to be located whereas, in this case, the node is already known, it is the 
     * CTBookmark itself. The only information that must be discovered is 
     * whether there is a run immediately in front of the boookmarkStart tag and 
     * whether that run is styled. If there is and if it is, then this style 
     * must be cloned and applied the text which will be inserted into the 
     * paragraph. 
     * 
     * @param bookmark An instance of the CTBookmark class that encapsulates 
     * information about the bookmark. 
     * @param run An instance of the XWPFRun class that encapsulates the text 
     * that is to be inserted into the document following the bookmark. 
     * @param para An instance of the XWPFParagraph class that encapsulates that 
     * part of the document, a paragraph, into which the run will be inserted. 
     */ 
    private void insertBeforeBookmark(CTBookmark bookmark, XWPFRun run, 
            XWPFParagraph para) { 
        Node insertBeforeNode = null; 
        Node childNode = null; 
        Node styleNode = null; 

        // Get the dom node from the bookmarkStart tag and look for another 
        // node immediately preceding it. 
        insertBeforeNode = bookmark.getDomNode(); 
        childNode = insertBeforeNode.getPreviousSibling(); 

        // If a node is found, try to get the styling from it. 
        if (childNode != null) { 
            styleNode = this.getStyleNode(childNode); 

            // If that previous node was styled, then apply this style to the 
            // text which will be inserted. 
            if (styleNode != null) { 
                run.getCTR().getDomNode().insertBefore( 
                        styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild()); 
            } 
        } 

        // Insert the text into the paragraph immediately in front of the 
        // bookmarkStart tag. 
        para.getCTP().getDomNode().insertBefore( 
                run.getCTR().getDomNode(), insertBeforeNode); 
    } 

    /** 
     * Replace the text - if any - contained between the bookmarkStart and it's 
     * matching bookmarkEnd tag with the text specified. The technique used will 
     * resemble that employed when inserting text after the bookmark. In short, 
     * the code will iterate along the nodes until it encounters a matching 
     * bookmarkEnd tag. Each node encountered will be deleted unless it is the 
     * final node before the bookmarkEnd tag is encountered and it is a 
     * character run. If this is the case, then it can simply be updated to 
     * contain the text the users wishes to see inserted into the document. If 
     * the last node is not a character run, then it will be deleted, a new run 
     * will be created and inserted into the paragraph between the bookmarkStart 
     * and bookmarkEnd tags. 
     * 
     * @param bookmark An instance of the CTBookmark class that encapsulates 
     * information about the bookmark. 
     * @param run An instance of the XWPFRun class that encapsulates the text 
     * that is to be inserted into the document following the bookmark. 
     * @param para An instance of the XWPFParagraph class that encapsulates that 
     * part of the document, a paragraph, into which the run will be inserted. 
     */ 
    private void replaceBookmark(CTBookmark bookmark, XWPFRun run, 
            XWPFParagraph para) { 
        Node nextNode = null; 
        Node styleNode = null; 
        Node lastRunNode = null; 
        NodeList childNodes = null; 
        Stack<Node> nodeStack = null; 
        boolean textNodeFound = false; 
        int bookmarkStartID = 0; 
        int bookmarkEndID = -1; 
        int numChildNodes = 0; 

        nodeStack = new Stack<Node>(); 
        bookmarkStartID = bookmark.getId().intValue(); 
        nextNode = bookmark.getDomNode(); 

        // Loop through the nodes looking for a matching bookmarkEnd tag 
        while (bookmarkStartID != bookmarkEndID) { 

            nextNode = nextNode.getNextSibling(); 

            // If an end tag is found, does it match the start tag? If so, end 
            // the while loop. 
            if (nextNode.getNodeName().contains("bookmarkEnd")) { 
                try { 
                    bookmarkEndID = Integer.parseInt( 
                            nextNode.getAttributes().getNamedItem("w:id").getNodeValue()); 
                } catch (NumberFormatException nfe) { 
                    bookmarkEndID = bookmarkStartID; 
                } 
            } else { 
                // If this is not a bookmark end tag, store the reference to the 
                // node on the stack for later deletion. This is easier that 
                // trying to delete the nodes as they are found. 
                nodeStack.push(nextNode); 
            } 
        } 

        // If the stack of nodes found between the bookmark tags is not empty 
        // then they have to be removed. 
        if (!nodeStack.isEmpty()) { 

            // Check the node at the top of the stack. If it is a run, get it's 
            // style - if any - and apply to the run that will be replacing it. 
            lastRunNode = nodeStack.pop(); 
            if ((lastRunNode.getNodeName().equals("w:r"))) { 
                styleNode = this.getStyleNode(lastRunNode); 
                if (styleNode != null) { 
                    run.getCTR().getDomNode().insertBefore( 
                            styleNode.cloneNode(true), run.getCTR().getDomNode().getFirstChild()); 
                } 
            } 

            // Delete any and all node that were found in between the start and 
            // end tags. This is slightly safer that trying to delete the nodes 
            // as they are found wile stepping through them in the loop above. 
            para.getCTP().getDomNode().removeChild(lastRunNode); 
            // Now, delete the remaing Nodes on the stack 
            while (!nodeStack.isEmpty()) { 
                para.getCTP().getDomNode().removeChild(nodeStack.pop()); 
            } 
        } 

        // Place the text into position, between the bookmark tags. 
        para.getCTP().getDomNode().insertBefore( 
                run.getCTR().getDomNode(), nextNode); 
    } 

    /** 
     * Recover styling information - if any - from another document node. Note 
     * that it is only possible to accomplish this if the node is a run (w:r) 
     * and this could be tested for in the code that calls this method. However, 
     * a check is made in the calling code as to whether a style has been found 
     * and only if a style is found is it applied. This method always returns 
     * null if it does nto find a style making that checking process easier. 
     * 
     * @param parentNode An instance of the Node class that encapsulates a 
     * reference to a document node. 
     * @return An instance of the Node class that encapsulates the styling 
     * information applied to a character run. Note that if no styling 
     * information is found in the run OR if the node passed as an argument to 
     * the parentNode parameter is NOT a run, then a null value will be 
     * returned. 
     */ 
    private Node getStyleNode(Node parentNode) { 
        Node childNode = null; 
        Node styleNode = null; 
        if (parentNode != null) { 

            // If the node represents a run and it has child nodes then 
            // it can be processed further. Note, whilst testing the code, it 
            // was observed that although it is possible to get a list of a nodes 
            // children, even when a node did have children, trying to obtain this 
            // list would often return a null value. This is the reason why the 
            // technique of stepping from one node to the next is used here. 
            if (parentNode.getNodeName().equalsIgnoreCase("w:r") 
                    && parentNode.hasChildNodes()) { 

                // Get the first node and catch it's reference for return if 
                // the first child node is a style node (w:rPr). 
                childNode = parentNode.getFirstChild(); 
                if (childNode.getNodeName().equals("w:rPr")) { 
                    styleNode = childNode; 
                } else { 
                    // If the first node was not a style node and there are other 
                    // child nodes remaining to be checked, then step through 
                    // the remaining child nodes until either a style node is 
                    // found or until all child nodes have been processed. 
                    while ((childNode = childNode.getNextSibling()) != null) { 
                        if (childNode.getNodeName().equals("w:rPr")) { 
                            styleNode = childNode; 
                            // Note setting to null here if a style node is 
                            // found in order order to terminate any further 
                            // checking 
                            childNode = null; 
                        } 
                    } 
                } 
            } 
        } 
        return (styleNode); 
    } 

    /** 
     * Recover and return any text that may exist within the document between 
     * the opening and closing brackets ([ and ]) of the named bookmark. 
     * 
     * @param bookmarkName An instance of the String class that encapsulates the 
     * name of the bookmark. 
     * @return An instance of the String class that encapsulates the text 
     * discovered between the opening and closing brackets (as seen when viewing 
     * the document with Word), if any. Note that a null value will be returned 
     * if the bookmark cannot be found. Also note that the code will look for 
     * bookmarks in the body of the document and individual table cells. 
     * @throws XmlException Thrown if a problem is encountered parsing the XML 
     * markup recovered from the document. 
     * @throws IOException Thrown if a problem is encountered within the 
     * underlying file system. 
     */ 
    public String getBookmarkText(String bookmarkName) throws XmlException, 
            IOException { 
        List<XWPFTable> tableList = null; 
        Iterator<XWPFTable> tableIter = null; 
        List<XWPFTableRow> rowList = null; 
        Iterator<XWPFTableRow> rowIter = null; 
        List<XWPFTableCell> cellList = null; 
        Iterator<XWPFTableCell> cellIter = null; 
        XWPFTable table = null; 
        XWPFTableRow row = null; 
        XWPFTableCell cell = null; 
        String text = null; 

        // Firstly, deal with any paragraphs in the body of the document. 
        text = this.procParasForBookmarkText(this.document.getParagraphs(), 
                bookmarkName); 

        // Then check to see if there are any bookmarks in table cells. To do this 
        // it is necessary to get at the list of paragraphs 'stored' within the 
        // individual table cell, hence this code which get the tables from the 
        // document, the rows from each table, the cells from each row and the 
        // paragraphs from each cell. 
        if (text == null) { 
            tableList = this.document.getTables(); 
            tableIter = tableList.iterator(); 
            while (tableIter.hasNext()) { 
                table = tableIter.next(); 
                rowList = table.getRows(); 
                rowIter = rowList.iterator(); 
                while (rowIter.hasNext()) { 
                    row = rowIter.next(); 
                    cellList = row.getTableCells(); 
                    cellIter = cellList.iterator(); 
                    while (cellIter.hasNext()) { 
                        cell = cellIter.next(); 
                        text = this.procParasForBookmarkText(cell.getParagraphs(), 
                                bookmarkName); 
                        if (text != null) { 
                            return (text); 
                        } 
                    } 
                } 
            } 
        } 
        return (text); 
    } 

    /** 
     * Processes a List of XWPFParagraph objects searching for the named 
     * bookmark. When the bookmark is found, any text that would appear between 
     * a bookmarks enclosing brackets ([ and ]) in the document as viewed using 
     * Word will actually be contained within one or more character run (w:r) 
     * elements that appear in the XML markup between the bookmarkStart and 
     * bookmarkEnd tags, a little like this; 
     * 
     * <pre>
     * 
     * <w:bookmarkStart w:id="3" w:name="InStyledText"/>
     *   <w:r w:rsidRPr="00DA438C">
     *     <w:rPr>
     *       <w:rFonts w:ascii="Engravers MT" w:hAnsi="Engravers MT" w:cs="Arimo"/>
     *       <w:color w:val="FF0000"/>
     *     </w:rPr>
     *     <w:t>
     *       text 
     *     </w:t>
     *   </w:r>
     * <w:bookmarkEnd w:id="3"/>
     * 
     * </pre>
     * 
     * which shows the markup for a bookmark called InStyledText. It has a 
     * single run that has a style applied to it and which contains a single 
     * piece of text. This text is held in a child node (w:t) and it is this 
     * child node (or these child nodes in case the run contains mode that one 
     * piece of text) that this code recovers. 
     * 
     * @param paraList A List containing one or more instances of the 
     * XWPFParagraph class. These are to be searched for the named bookmark. 
     * @param bookmarkName An instance of the String class that encapsulates the 
     * name of the bookmark. 
     * @return An instance of the String class encapsulating the text - if any - 
     * found between the bookmarks start and end tags. A null value will be 
     * returned if the bookmark cannot be found. 
     * @throws XmlException Thrown if a problem is encountered parsing the XML 
     * markup recovered from the document in order to construct a CTText 
     * instance which is required to obtain the bookmarks text. 
     * @throws IOException An OutputStream is used to read the contents of the 
     * CTText object and an IOException will be thrown if any problems are 
     * encountered. 
     */ 
    public String procParasForBookmarkText(List<XWPFParagraph> paraList, 
            String bookmarkName) throws XmlException, IOException { 
        Iterator<XWPFParagraph> paraIter = null; 
        XWPFParagraph para = null; 
        XWPFRun run = null; 
        List<CTBookmark> bookmarkList = null; 
        Iterator<CTBookmark> bookmarkIter = null; 
        CTBookmark bookmark = null; 
        StringBuilder builder = null; 

        // Get an Iterator to step through the contents of the paragraph list. 
        paraIter = paraList.iterator(); 
        while (paraIter.hasNext()) { 

            // Get the paragraph, a llist of CTBookmark objects and an Iterator 
            // to step through the list of CTBookmarks. 
            para = paraIter.next(); 
            bookmarkList = para.getCTP().getBookmarkStartList(); 
            bookmarkIter = bookmarkList.iterator(); 
            while (bookmarkIter.hasNext()) { 

                // Get a Bookmark and check it's name. If the name of the 
                // bookmark matches the name the user has specified then get the 
                // bookmarks ID. This is required to cope with the situation where 
                // one bookmark is nested within another; we do not want to end 
                // processing until we hit the matching bookmarkEnd tag. 
                bookmark = bookmarkIter.next(); 
                if (bookmark.getName().equals(bookmarkName)) { 
                    builder = this.getTextFromBookmark(bookmark); 
                } 
            } 
        } 
        return (builder == null ? null : builder.toString()); 
    } 

    /** 
     * There are two types of bookmarks. One is a simple placeholder whilst the 
     * second is still a placeholder but it 'contains' some text. In the second 
     * instance, the creator of the document has selected some text and then 
     * chosen to insert a bookmark there and the difference if obvious when 
     * looking at the XML markup. 
     * 
     * The simple case; 
     * 
     * <pre>
     * 
     * <w:bookmarkStart w:name="AllAlone" w:id="0"/><w:bookmarkEnd w:id="0"/>
     * 
     * </pre>
     * 
     * The more complex case; 
     * 
     * <pre>
     * 
     * <w:bookmarkStart w:name="InStyledText" w:id="3"/>
     *   <w:r w:rsidRPr="00DA438C">
     *     <w:rPr>
     *       <w:rFonts w:hAnsi="Engravers MT" w:ascii="Engravers MT" w:cs="Arimo"/>
     *       <w:color w:val="FF0000"/>
     *     </w:rPr>
     *     <w:t>text</w:t>
     *   </w:r>
     * <w:bookmarkEnd w:id="3"/>
     * 
     * </pre>
     * 
     * This method assumes that the user wishes to recover the content from any 
     * character run that appears in the markup between a matching pair of 
     * bookmarkStart and bookmarkEnd tags; thus, using the example above again, 
     * this method would return the String 'text' to the user. It is possible 
     * however for a bookmark to contain more than one run and for a bookmark to 
     * contain other bookmarks. In both of these cases, this code will return 
     * the text contained within any and all runs that appear in the XML markup 
     * between matching bookmarkStart and bookmarkEnd tags. The term 'matching 
     * bookmarkStart and bookmarkEndtags' here means tags whose id attributes 
     * have matching value. 
     * 
     * @param bookmark An instance of the CTBookmark class that encapsulates 
     * information about a bookmark in a Word document. 
     * @return An instance of the StringBuilder class encapsulating the text 
     * recovered from any character run elements found between the bookmark's 
     * start and end tags. If no text is found then a null value will be 
     * returned. 
     * @throws XmlException Thrown if a problem is encountered parsing the XML 
     * markup recovered from the document in order to construct a CTText 
     * instance which is required to obtain the bookmarks text. 
     * @throws IOException An OutputStream is used to read the contents of the 
     * CTText object and an IOException will be thrown if any problems are 
     * encountered. 
     */ 
    private StringBuilder getTextFromBookmark(CTBookmark bookmark) 
            throws IOException, XmlException { 
        int startBookmarkID = 0; 
        int endBookmarkID = -1; 
        Node nextNode = null; 
        Node childNode = null; 
        CTText text = null; 
        ByteArrayOutputStream baos = null; 
        StringBuilder builder = null; 
        String rawXML = null; 

        // Get the ID of the bookmark from it's start tag, the DOM node from the 
        // bookmark (to make looping easier) and initialise the StringBuilder. 
        startBookmarkID = bookmark.getId().intValue(); 
        nextNode = bookmark.getDomNode(); 
        builder = new StringBuilder(); 

        // Loop through the nodes held between the bookmark's start and end 
        // tags. 
        while (startBookmarkID != endBookmarkID) { 

            // Get the next node and, if it is a bookmarkEnd tag, get it's ID 
            // as matching ids will terminate the while loop.. 
            nextNode = nextNode.getNextSibling(); 
            if (nextNode.getNodeName().contains("bookmarkEnd")) { 

                // Get the ID attribute from the node. It is a String that must 
                // be converted into an int. An exception could be thrown and so 
                // the catch clause will ensure the loop ends neatly even if the 
                // value might be incorrect. Must inform the user. 
                try { 
                    endBookmarkID = Integer.parseInt( 
                            nextNode.getAttributes(). 
                            getNamedItem("w:id").getNodeValue()); 
                } catch (NumberFormatException nfe) { 
                    endBookmarkID = startBookmarkID; 
                } 
            } else { 
                // This is not a bookmarkEnd node and can processed it for any 
                // text it may contain. Note the check for both type - it must 
                // be a run - and contain children. Interestingly, it seems as 
                // though the node may contain children and yet the call to 
                // nextNode.getChildNodes() will still return an empty list, 
                // hence the need to step through the child nodes. 
                if (nextNode.getNodeName().equals("w:r") 
                        && nextNode.hasChildNodes()) { 
                    // Get the text from the child nodes. 
                    builder.append(this.getTextFromChildNodes(nextNode)); 
                } 
            } 
        } 
        return (builder); 
    } 

    /** 
     * Iterates through all and any children of the Node whose reference will be 
     * passed as an argument to the node parameter, and recover the contents of 
     * any text nodes. Testing revealed that a node can be called a text node 
     * and yet report it's type as being something different, an element node 
     * for example. Calling the getNodeValue() method on a text node will return 
     * the text the node encapsulates but doing the same on an element node will 
     * not. In fact, the call will simply return a null value. As a result, this 
     * method will test the nodes name to catch all text nodes - those whose 
     * name is to 'w:t' and then it's type. If the type is reported to be a text 
     * node, it is a trivial task to get at it's contents. However, if the type 
     * is not reported as a text type, then it is necessary to parse the raw XML 
     * markup for the node to recover it's value. 
     * 
     * @param node An instance of the Node class that encapsulates a reference 
     * to a node recovered from the document being processed. It should be 
     * passed a reference to a character run - 'w:r' - node. 
     * @return An instance of the String class that encapsulates the text 
     * recovered from the nodes children, if they are text nodes. 
     * @throws IOException Thrown if a problem occurs in the underlying file 
     * system and only necessary as a stream may be used to recover the raw XML 
     * markup for a child node. 
     * @throws XmlException Thrown if a problem is encountered parsing a nodes 
     * raw XML markup in order to construct a openxml4j CTText object. 
     */ 
    private String getTextFromChildNodes(Node node) throws IOException, 
            XmlException { 
        NodeList childNodes = null; 
        Node childNode = null; 
        CTText text = null; 
        StringBuilder builder = new StringBuilder(); 
        int numChildNodes = 0; 

        // Get a list of chid nodes from the node passed to the method and 
        // find out how many children there are in the list. 
        childNodes = node.getChildNodes(); 
        numChildNodes = childNodes.getLength(); 

        // Iterate through the children one at a time - it is possible for a 
        // run to ciontain zero, one or more text nodes - and recover the text 
        // from an text type child nodes. 
        for (int i = 0; i < numChildNodes; i++) { 

            // Get a node and check it's name. If this is 'w:t' then process as 
            // text type node. 
            childNode = childNodes.item(i); 

            if (childNode.getNodeName().equals("w:t")) { 

                // If the node reports it's type as txet, then simply call the 
                // getNodeValue() method to get at it's text. 
                if (childNode.getNodeType() == Node.TEXT_NODE) { 
                    builder.append(childNode.getNodeValue()); 
                } else { 
                    // Correct the type by parsing the node's XML markup and 
                    // creating a CTText object. Call the getStringValue() 
                    // method on that to get the text. 
                    text = CTText.Factory.parse(childNode); 
                    builder.append(text.getStringValue()); 
                } 
            } 
        } 
        return (builder.toString()); 
    } 

    public static void main(String[] args) { 
        try { 
            // open the existing workbook and get the text of two bookmarks. 
            DOCXTest docxTest = new DOCXTest(); 
            docxTest.openFile("D:/xiezuoweit-v1.docx");
            System.out.println("andy contains: " + docxTest.getBookmarkText("andy")); 
            docxTest.insertAtBookmark("andy", "This should replace the EDMS_Bookmark1", DOCXTest.REPLACE); 
            docxTest.saveAs("D:/xiezuoweit-v1.docx");
            /*docxTest.openFile("C:/temp/BeforeBookMarkValuesReplacement.docx"); 
            System.out.println("EDMS_Bookmark1 contains: " + docxTest.getBookmarkText("EDMS_Bookmark1")); 
            System.out.println("EDMS_Bookmark5 contains: " + docxTest.getBookmarkText("EDMS_Bookmark5")); 
            // Replace the text at those two bookmakrs and then save the file away 
            // using a different name 
            docxTest.insertAtBookmark("EDMS_Bookmark1", "This should replace the EDMS_Bookmark1", DOCXTest.REPLACE); 
            docxTest.insertAtBookmark("EDMS_Bookmark5", "This should replace the EDMS_Bookmark5", DOCXTest.REPLACE); 
            docxTest.saveAs("C:/temp/AfterBookMarkValuesReplacement.docx"); 
            // Open the new file and demonstrate that the bookamrk text has changed. 
            docxTest.openFile("C:/temp/AfterBookMarkValuesReplacement.docx"); 
            System.out.println("EDMS_Bookmark1 contains: " + docxTest.getBookmarkText("EDMS_Bookmark1")); 
            System.out.println("EDMS_Bookmark5 contains: " + docxTest.getBookmarkText("EDMS_Bookmark5")); */
            
        } catch (Exception ex) { 
            System.out.println("Caught a: " + ex.getClass().getName()); 
            System.out.println("Message: " + ex.getMessage()); 
            System.out.println("Stacktrace follows:....."); 
            ex.printStackTrace(System.out); 
        } 
    } 
}
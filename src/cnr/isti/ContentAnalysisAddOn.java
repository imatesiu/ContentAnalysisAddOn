package cnr.isti;

import cnr.isti.xml.data.Annotation;
import cnr.isti.xml.data.QualityCriteria;
import cnr.isti.xml.data.collaborative.AnnotatedCollaborativeContentAnalyses;
import cnr.isti.xml.data.collaborative.AnnotatedCollaborativeContentAnalysis;
import cnr.isti.xml.data.collaborative.CollaborativeContent;
import cnr.isti.xml.data.collaborative.CollaborativeContentAnalysis;
import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyState;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceDisplayName;
import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.linguistic2.XLinguServiceEventBroadcaster;
import com.sun.star.linguistic2.XLinguServiceEventListener;
import com.sun.star.linguistic2.XProofreader;
import com.sun.star.table.BorderLine;
import com.sun.star.table.TableBorder;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XTableRows;
import com.sun.star.task.XJobExecutor;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XWordCursor;
import java.awt.Color;
import java.awt.Rectangle;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import javax.swing.JOptionPane;
import javax.ws.rs.client.Client;
import javax.ws.rs.client.ClientBuilder;
import javax.ws.rs.client.Entity;
import javax.ws.rs.client.WebTarget;
import javax.ws.rs.core.GenericType;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

public final class ContentAnalysisAddOn extends WeakBase
        implements com.sun.star.frame.XDispatchProvider,
        com.sun.star.frame.XDispatch, com.sun.star.lang.XInitialization,
        com.sun.star.lang.XServiceInfo, XProofreader, XLinguServiceEventBroadcaster, XServiceDisplayName, XJobExecutor {

    private final XComponentContext m_xContext;
    private com.sun.star.frame.XFrame m_xFrame;
    private static final String m_implementationName = ContentAnalysisAddOn.class.getName();
    private static final String[] m_serviceNames = {
        "com.sun.star.linguistic2.Proofreader",
        "com.sun.star.frame.ProtocolHandler"};

    private XNameContainer xNameCont = null;

    private final List<XLinguServiceEventListener> xEventListeners;
    private boolean recheck;
    private static final boolean testMode = true;
    private static final QualityCriteria qc = new QualityCriteria();

    @Override
    public String getServiceDisplayName(Locale locale) {
        return "LanguageCATool";
    }

    @Override
    public void trigger(String sEvent) {
        if (Thread.currentThread().getContextClassLoader() == null) {
            Thread.currentThread().setContextClassLoader(ContentAnalysisAddOn.class.getClassLoader());
        }
        /*if (!javaVersionOkay()) {
	      return;
	    }*/
        try {
            if ("configure".equals(sEvent)) {
                //  runOptionsDialog();
            } else if ("about".equals(sEvent)) {
                // AboutDialogThread aboutThread = new AboutDialogThread(MESSAGES);
                //  aboutThread.start();
            } else {
                System.err.println("Sorry, don't know what to do, sEvent = " + sEvent);
            }
        } catch (Throwable e) {
            showError(e);
        }
    }

    @Override
    public final boolean addLinguServiceEventListener(XLinguServiceEventListener eventListener) {
        if (eventListener == null) {
            return false;
        }
        xEventListeners.add(eventListener);
        return true;
    }

    /**
     * Remove a listener from the event listeners list.
     *
     * @param eventListener the listener to be removed
     * @return true if listener is non-null and has been removed, false
     * otherwise
     */
    @Override
    public final boolean removeLinguServiceEventListener(XLinguServiceEventListener eventListener) {
        if (eventListener == null) {
            return false;
        }
        if (xEventListeners.contains(eventListener)) {
            xEventListeners.remove(eventListener);
            return true;
        }
        return false;
    }

    public boolean isSpellChecker() {
        return false;
    }

    public Locale[] getLocales() {
        List<Locale> locales = new ArrayList<Locale>();
        locales.add(new Locale("it-It", "", ""));
        locales.add(new Locale("en-US", "", ""));
      //  locales.add(new Locale("en-GB", "", ""));
        locales.add(new Locale("en-UK", "", ""));
        return locales.toArray(new Locale[locales.size()]);
    }
  
    public boolean hasLocale(Locale arg0) {
        return true;
    }

    public void ignoreRule(String arg0, Locale arg1) {
    }

    public void resetIgnoreRules() {
        recheck = true;
    }

    public ProofreadingResult doProofreading(String docID,
            String paraText, Locale locale, int startOfSentencePos,
            int nSuggestedBehindEndOfSentencePosition,
            PropertyValue[] propertyValues) {
        ProofreadingResult paRes = new ProofreadingResult();
        try {
            paRes.nStartOfSentencePosition = startOfSentencePos;
            paRes.xProofreader = this;
            paRes.aLocale = locale;
            paRes.aDocumentIdentifier = docID;
            paRes.aText = paraText;
            paRes.aProperties = propertyValues;
            int[] footnotePositions = getPropertyValues("FootnotePositions", propertyValues);  // since LO 4.3
            return doGrammarCheckingInternal(paraText, locale, paRes, footnotePositions, nSuggestedBehindEndOfSentencePosition);
        } catch (Throwable t) {
            showError(t);
            return paRes;
        }

    }

    private int[] getPropertyValues(String propName, PropertyValue[] propertyValues) {
        for (PropertyValue propertyValue : propertyValues) {
            if (propName.equals(propertyValue.Name)) {
                if (propertyValue.Value instanceof int[]) {
                    return (int[]) propertyValue.Value;
                } else {
                    System.err.println("Not of expected type int[]: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
                }
            }
        }
        return new int[]{};  // e.g. for LO/OO < 4.3 and the 'FootnotePositions' property
    }

    private synchronized ProofreadingResult doGrammarCheckingInternal(
            String paraText, Locale locale, ProofreadingResult paRes, int[] footnotePositions, int nSuggestedBehindEndOfSentencePosition) {

        try {
            if (paraText.length() > 1) {

                AnnotatedCollaborativeContentAnalyses res = CheckText(paraText);
                if (res != null) {
                    fromCAtoProofreadingResult(res, paRes, paraText, footnotePositions);
                    System.out.println(res);
                }
            }
        } catch (Throwable t) {

            showError(t);
            paRes.nBehindEndOfSentencePosition = paraText.length();
        }

        return paRes;
    }

    private AnnotatedCollaborativeContentAnalyses CheckText(String paraText) {
        try {
            if (paraText.length() > 1) {
                Client client = ClientBuilder.newClient();
                WebTarget target = client.target("http://contentanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent");

                CollaborativeContentAnalysis cca = new CollaborativeContentAnalysis();
                cca.setLanguage(this.getLanguage());

                cca.setCollaborativeContent(new CollaborativeContent(String.valueOf(this.getId()), this.getTitle()));
                cca.getCollaborativeContent().setContentplain(paraText);
                cca.setTypeofdoc("OoOAddon");
                cca.setQualityCriteria(qc);

                Entity<CollaborativeContentAnalysis> entity = Entity.entity(cca, MediaType.APPLICATION_XML);
                //GenericEntity<JAXBElement<CollaborativeContentAnalysis>> gw = new GenericEntity<JAXBElement<CollaborativeContentAnalysis>>(cca){};
                Response response = target.request(MediaType.APPLICATION_XML).post(entity);

                String id = response.readEntity(String.class);

                client = ClientBuilder.newClient();
                if (id == null) {
                    return null;
                }

                target = client.target("http://contentanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent/" + id + "/status");
                String status = "";
                while (!status.equals("OK")) {

                    status = target.request().get(String.class);

                    //this.setStatus(status);
                    if (status.equals("ERROR")) {
                        break;
                    }
                }
                //log.trace("Status: "+status);

                if (status.equals("OK")) {

                    String ipAddress = null;
                    if (ipAddress == null) {
                        ipAddress="0.0.0.0";
                    }
                    //System.out.println("ipAddress:" + ipAddress);

                    target = client.target("http://contentanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent/" + id);
                    Response annotatecontent = target.request().header("X-FORWARDED-FOR", ipAddress).get();
                    AnnotatedCollaborativeContentAnalyses res = annotatecontent.readEntity(new GenericType<AnnotatedCollaborativeContentAnalyses>() {
                    });
                    //this.setCollectionannotatedcontent(res.getAnnotateCollaborativeContentAnalysis());
                    return res;
                }
            }
        } catch (Throwable t) {
            showError(t);
            return null;
            //todo
        }
        return null;
    }

    private void fromCAtoProofreadingResult(AnnotatedCollaborativeContentAnalyses res, ProofreadingResult paRes,
            String paraText, int[] footnotePositions) {
        List<SingleProofreadingError> errorList = new ArrayList<SingleProofreadingError>();
        List<Annotation> listanna = new ArrayList<Annotation>();
        for (AnnotatedCollaborativeContentAnalysis acca : res.getAnnotateCollaborativeContentAnalysis()) {
            listanna.addAll(acca.getAnnotations());
        }
        //Collections.sort(listanna);
        for (Annotation anna : listanna) {
            errorList.add(createOOoError(anna));
        }
        if (!errorList.isEmpty()) {
            SingleProofreadingError[] errorArray = errorList.toArray(new SingleProofreadingError[errorList.size()]);
            Arrays.sort(errorArray, new ErrorPositionComparator());
            paRes.aErrors = errorArray;
        }

    }

    /**
     * Creates a SingleGrammarError object for use in LO/OO.
     */
    private SingleProofreadingError createOOoError(Annotation aa) {
        SingleProofreadingError aError = new SingleProofreadingError();
        aError.nErrorType = TextMarkupType.PROOFREADING;
        // the API currently has no support for formatting text in comments
        aError.aFullComment = aa.getRecommendation().replaceAll("<suggestion>", "\"").replaceAll("</suggestion>", "\"")
                .replaceAll("([\r]*\n)", " ");
        // not all rules have short comments

        aError.aShortComment = aa.getType();

        int numSuggestions = 1;
        String[] allSuggestions = new String[numSuggestions];
        allSuggestions[0] = aa.getRecommendation();

        aError.aSuggestions = allSuggestions;
        Integer startoff = aa.getstartNode_Offset();
        Integer endoff = aa.getendNode_Offset();
        aError.nErrorStart = startoff;
        aError.nErrorLength = endoff - startoff;
        aError.aRuleIdentifier = aa.getId().toString();

        // LibreOffice since version 3.5 supports an URL that provides more
        // information about the error,
        // older version will simply ignore the property:
        //if (ruleMatch.getRule().getUrl() != null) {
        //   aError.aProperties = new PropertyValue[] { new PropertyValue(
        //       "FullCommentURL", -1, ruleMatch.getRule().getUrl().toString(),
        //      PropertyState.DIRECT_VALUE) };
        //} else {
        aError.aProperties = new PropertyValue[0];
        //}
        return aError;
    }

    public ContentAnalysisAddOn(XComponentContext context) {
        m_xContext = context;
        recheck = false;
        xEventListeners = new ArrayList<XLinguServiceEventListener>();
        //qc = new QualityCriteria();
        qc.setCorrectness(true);
        qc.setSimplicity(true);
        qc.setContentClarity(true);
        qc.setNonAmbiguity(true);
        qc.setCompleteness(false);
        qc.setPresentationClarity(false);
    }

    ;

    public static XSingleComponentFactory __getComponentFactory(String sImplementationName) {
        XSingleComponentFactory xFactory = null;

        if (sImplementationName.equals(m_implementationName)) {
            xFactory = Factory.createComponentFactory(ContentAnalysisAddOn.class, m_serviceNames);
        }
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo(XRegistryKey xRegistryKey) {
        return Factory.writeRegistryServiceInfo(m_implementationName,
                m_serviceNames,
                xRegistryKey);
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch queryDispatch(com.sun.star.util.URL aURL,
            String sTargetFrameName,
            int iSearchFlags) {
        if (aURL.Protocol.compareTo("cnr.isti.contentanalysisaddon:") == 0) {
            if (aURL.Path.compareTo("Config") == 0) {
                return this;
            }
            if (aURL.Path.compareTo("Report") == 0) {
                return this;
            }

            if (aURL.Path.compareTo("About") == 0) {
                return this;
            }
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch[] queryDispatches(
            com.sun.star.frame.DispatchDescriptor[] seqDescriptors) {
        int nCount = seqDescriptors.length;
        com.sun.star.frame.XDispatch[] seqDispatcher
                = new com.sun.star.frame.XDispatch[seqDescriptors.length];

        for (int i = 0; i < nCount; ++i) {
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,
                    seqDescriptors[i].FrameName,
                    seqDescriptors[i].SearchFlags);
        }
        return seqDispatcher;
    }

    // com.sun.star.frame.XDispatch:
    public void dispatch(com.sun.star.util.URL aURL,
            com.sun.star.beans.PropertyValue[] aArguments) {
        if (aURL.Protocol.compareTo("cnr.isti.contentanalysisaddon:") == 0) {
            if (aURL.Path.compareTo("Config") == 0) {
                // add your own code here
                System.out.println(aURL.Path);
                try {
                    createDialog();
                } catch (Exception e) {
                    showError(e);
                    throw new com.sun.star.lang.WrappedTargetRuntimeException(e.getMessage(), this, e);
                }
                return;
            }

            if (aURL.Path.compareTo("Report") == 0) {
                // add your own code here
                System.out.println(aURL.Path);
                try {
                    createReport();
                } catch (Exception e) {
                    showError(e);
                    throw new com.sun.star.lang.WrappedTargetRuntimeException(e.getMessage(), this, e);
                }
                return;
            }

            if (aURL.Path.compareTo("About") == 0) {
                // add your own code here
                System.out.println(aURL.Path);
                try {
                    createAboutDialog();
                } catch (Exception e) {
                    showError(e);
                    throw new com.sun.star.lang.WrappedTargetRuntimeException(e.getMessage(), this, e);
                }
                return;
            }

        }
    }
    
    public static void insertHeadersIntoCell(String sCellName, String sText,
        XTextTable xTable) throws UnknownPropertyException, PropertyVetoException,
        com.sun.star.lang.IllegalArgumentException, WrappedTargetException, IndexOutOfBoundsException {

    // Access the XText interface of the cell referred to by sCellName (e.g., A1):
    XText xCellText = (XText) UnoRuntime.queryInterface(XText.class, xTable.getCellByName(sCellName));

    // Set the text in the cell to sText (e.g., "Nick Name"):
    xCellText.setString(sText);

    //BACKGROUND COLORS:
    // Select the table headers and get the cell properties:
    XCellRange xCellRange = ( XCellRange )UnoRuntime.queryInterface( XCellRange.class, xTable );
    XCellRange xSelectedCells = xCellRange.getCellRangeByName("A1:E1");
    XPropertySet xCellProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xSelectedCells);
        // Format the color of the table headers (page 56 and 57):
    XPropertySet xTableProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xCellRange);
    xCellProps.setPropertyValue("BackColor", new Integer(0x000052));

    // BORDERS:
    // Define a border line, then assign a color and a width:
    BorderLine theLine = new BorderLine();
    theLine.Color = 0x000099;
    theLine.OuterLineWidth = 1;
    // Apply the line definition to all cell borders and make them valid:
    TableBorder bord = new TableBorder();
    bord.VerticalLine = bord.HorizontalLine =
            bord.LeftLine = bord.RightLine =
            bord.TopLine = bord.BottomLine =
            theLine;
    bord.IsVerticalLineValid = bord.IsHorizontalLineValid =
            bord.IsLeftLineValid = bord.IsRightLineValid =
            bord.IsTopLineValid = bord.IsBottomLineValid =
            true;
    xTableProps.setPropertyValue("TableBorder", bord);

}

    private synchronized void createReport() {
        try {

            XTextDocument m_xTextDocument = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, m_xFrame.getController().getModel());
            String text = m_xTextDocument.getText().getString();
            String filename = m_xFrame.getController().getModel().getURL();
            if (text != null) {
                if (text.length() > 1) {
                    AnnotatedCollaborativeContentAnalyses TotalACA = new AnnotatedCollaborativeContentAnalyses();//CheckText(text);

                 //   if (!TotalACA.getAnnotateCollaborativeContentAnalysis().isEmpty()) {
                        com.sun.star.lang.XMultiComponentFactory xMCF = m_xContext.getServiceManager();
                        Object oDesktop = xMCF.createInstanceWithContext(
                                "com.sun.star.frame.Desktop", m_xContext);
                        com.sun.star.frame.XComponentLoader xCompLoader
                                = (com.sun.star.frame.XComponentLoader) UnoRuntime.queryInterface(
                                        com.sun.star.frame.XComponentLoader.class, oDesktop);
                        String url = "private:factory/swriter";

                        PropertyValue[] propVals = new PropertyValue[1];
                        propVals[0] = new PropertyValue();
                        if(filename == null){
                            filename = "_blank";
                        }
                        propVals[0].Name = "Content Analysis Report of "+filename;
                        propVals[0].State = PropertyState.DIRECT_VALUE;
                        propVals[0].Handle = -1;
                        // propVals[0].Value = new uno.Any(true); // writer_pdf_Export  ,  swriter: MS Word 97 , HTML (StarWriter) ,*/

                        // PropertyValue propertyValue[] = new PropertyValue[2];
                        XComponent oDocToStore = xCompLoader.loadComponentFromURL(
                                url, "_blank", 0, propVals);
                        /*com.sun.star.frame.XStorable xStorable
                    = (com.sun.star.frame.XStorable) UnoRuntime.queryInterface(
                            com.sun.star.frame.XStorable.class, oDocToStore);*/
                        XTextDocument xTextDocument
                                = (XTextDocument) UnoRuntime.queryInterface(
                                        XTextDocument.class, oDocToStore);

                        insertTableReport(xTextDocument, TotalACA, text);

                        // xTextDocument.getText().setString(docText);
                        // 
                        /* propertyValue[0] = new com.sun.star.beans.PropertyValue();
            propertyValue[0].Name = "Overwrite";
            propertyValue[0].Value = new Boolean(true);
            propertyValue[1] = new com.sun.star.beans.PropertyValue();
            propertyValue[1].Name = "FilterName";
            propertyValue[1].Value = "StarOffice XML (Writer)";
            xStorable.storeAsURL( sSaveUrl.toString(), propertyValue );*/
               //     }
                }
            }
        } catch (Exception e) {
            showError(e);
            System.out.println(e.getMessage());
        }

    }

    
      private void insertTableReport(XTextDocument xTextDocument, AnnotatedCollaborativeContentAnalyses TotalACA, String text) throws Exception {
       // if (!TotalACA.getAnnotateCollaborativeContentAnalysis().isEmpty()) {
            com.sun.star.text.XText xText = xTextDocument.getText();
             // get internal service factory of the document
              XMultiServiceFactory xWriterFactory = (XMultiServiceFactory)UnoRuntime.queryInterface(
                  XMultiServiceFactory.class, xTextDocument);
 
              // insert TextTable and get cell text, then manipulate text in cell
              Object table = xWriterFactory.createInstance("com.sun.star.text.TextTable");
              
             /* XTextContent xTextContentTable = (XTextContent)UnoRuntime.queryInterface(
                 XTextContent.class, table);*/
              
              
              XTextTable xTextTable = (XTextTable)UnoRuntime.queryInterface(
                  XTextTable.class, table);
                 
            
              xTextTable.initialize(5, 5);
              xText.insertTextContent(xText.getEnd(), xTextTable, false);
             insertHeadersIntoCell("A1","Defect Number",xTextTable);
              insertHeadersIntoCell("B1","Defect",xTextTable);
               insertHeadersIntoCell("C1","Description",xTextTable);
             insertHeadersIntoCell("D1","Orginal Text",xTextTable);
              insertHeadersIntoCell("E1","FeedBack",xTextTable);
              
              
              
             
 
          /*    XCellRange xCellRange = (XCellRange)UnoRuntime.queryInterface(
                  XCellRange.class, table);
             XCell xCell = xCellRange.getCellByPosition(0, 1);
              XText xCellText = (XText)UnoRuntime.queryInterface(XText.class, xCell);
              xCellText.setString("Ciao");*/
              
        // Change properties of the range.
        // Accessing a cell range over its position.
//        XCellRange xCellRangeIntestazione = xCellRange.getCellRangeByName("A1:D1");
 //XPropertySet xPropSet = (com.sun.star.beans.XPropertySet)
 //    UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xCellRangeIntestazione);
 //xPropSet.setPropertyValue("CellBackColor", new Integer(0x8080FF));
             // manipulateText(xCellText);
            
       // }
      }
    private void insertReport(XTextDocument xTextDocument, AnnotatedCollaborativeContentAnalyses TotalACA, String text) throws Exception {
        if (!TotalACA.getAnnotateCollaborativeContentAnalysis().isEmpty()) {
            com.sun.star.text.XText xText = xTextDocument.getText();
            for (AnnotatedCollaborativeContentAnalysis acc : TotalACA.getAnnotateCollaborativeContentAnalysis()) {
                /* if(acc.getType().equals(acc)){
                    
                }*/
                String Text = acc.getType() + " Risk:\n\r";
                Text += acc.getOverallQualityMeasure() + " " + acc.getOverallQuality() + " - " + acc.getOverallRecommendations() + "\n\r";

                //xText.setString(Text);
                XTextRange pos = xText.getEnd();
                xText.insertString(pos, Text, true);
                XWordCursor xWordCursor
                        = (com.sun.star.text.XWordCursor) UnoRuntime.queryInterface(
                                com.sun.star.text.XWordCursor.class, pos);

                // xWordCursor.gotoStart(true);
                // xWordCursor.gotoEnd(true);
                // xWordCursor.gotoEndOfWord(true);
                XPropertySet xPropertySet
                        = (XPropertySet) UnoRuntime.queryInterface(
                                XPropertySet.class, xWordCursor);
                xPropertySet.setPropertyValue("CharWeight",
                        new Float(com.sun.star.awt.FontWeight.BOLD));
                xPropertySet.setPropertyValue("CharPosture", com.sun.star.awt.FontSlant.NONE);
                xPropertySet.setPropertyValue("CharBackColor", new java.lang.Integer(Color.WHITE.getRGB()));
                for (Annotation ann : acc.getAnnotations()) {

                    String raccomandation = ann.getRecommendation();

                    if (ann.getStartSentence_Offset() != null) {
                        Integer s = ann.getStartSentence_Offset();
                        Integer e = ann.getEndSentence_Offset();
                        String sentence = "";
                        try{
                            sentence = text.substring(s, e);
                        }catch(Exception error){
                            showError(error);
                            sentence= "not valid";
                        }
                        pos = xText.getEnd();
                        xText.insertString(pos, sentence+"\n\r", true);
                        xWordCursor
                                = (com.sun.star.text.XWordCursor) UnoRuntime.queryInterface(
                                        com.sun.star.text.XWordCursor.class, pos);
                        xPropertySet
                                = (XPropertySet) UnoRuntime.queryInterface(
                                        XPropertySet.class, xWordCursor);
                        xPropertySet.setPropertyValue("CharWeight",
                                new Float(com.sun.star.awt.FontWeight.DONTKNOW));
                        xPropertySet.setPropertyValue("CharPosture", com.sun.star.awt.FontSlant.NONE);
                        xPropertySet.setPropertyValue("CharBackColor", new java.lang.Integer(Color.WHITE.getRGB()));
                    }
                    // xText.insertString(xText, Text, recheck);
                    XTextRange pos2 = xText.getEnd();
                    xText.insertString(pos2, raccomandation + "\n\r", true);
                    XWordCursor xWordCursor2 = (XWordCursor) UnoRuntime.queryInterface(
                            XWordCursor.class, pos2);

                    XPropertySet xPropertySet2 = (XPropertySet) UnoRuntime.queryInterface(
                            XPropertySet.class, xWordCursor2);

                    xPropertySet2.setPropertyValue("CharPosture", com.sun.star.awt.FontSlant.ITALIC);
                    xPropertySet2.setPropertyValue("CharWeight",
                            new Float(com.sun.star.awt.FontWeight.DONTKNOW));
                    xPropertySet2.setPropertyValue("CharBackColor", 0x00FFFF00);

                }

            }
        }
    }

    private void createDialog() throws com.sun.star.uno.Exception {
        // get the service manager from the component context
        XMultiComponentFactory xMultiComponentFactory = m_xContext.getServiceManager();
        // create the dialog model and set the properties
        Object dialogModel = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialogModel", m_xContext);
        XPropertySet xPSetDialog = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, dialogModel);
        xPSetDialog.setPropertyValue("PositionX", new Integer(100));
        xPSetDialog.setPropertyValue("PositionY", new Integer(100));
        xPSetDialog.setPropertyValue("Width", new Integer(150));
        xPSetDialog.setPropertyValue("Height", new Integer(100));
        xPSetDialog.setPropertyValue("Title", new String("Runtime Dialog CA Configuration"));
        // get the service manager from the dialog model
        XMultiServiceFactory xMultiServiceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(
                XMultiServiceFactory.class, dialogModel);

        // create the button model and set the properties
        Object buttonModel = xMultiServiceFactory.createInstance(
                "com.sun.star.awt.UnoControlButtonModel");
        XPropertySet xPSetButton = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, buttonModel);
        xPSetButton.setPropertyValue("PositionX", new Integer(90));
        xPSetButton.setPropertyValue("PositionY", new Integer(80));
        xPSetButton.setPropertyValue("Width", new Integer(50));
        xPSetButton.setPropertyValue("Height", new Integer(14));
        xPSetButton.setPropertyValue("Name", "INVIO");
        xPSetButton.setPropertyValue("TabIndex", new Short((short) 0));
        xPSetButton.setPropertyValue("Label", new String("OK"));

        Object buttonModel2 = xMultiServiceFactory.createInstance(
                "com.sun.star.awt.UnoControlButtonModel");
        XPropertySet xPSetButton2 = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, buttonModel2);
        xPSetButton2.setPropertyValue("PositionX", new Integer(5));
        xPSetButton2.setPropertyValue("PositionY", new Integer(80));
        xPSetButton2.setPropertyValue("Width", new Integer(50));
        xPSetButton2.setPropertyValue("Height", new Integer(14));
        xPSetButton2.setPropertyValue("Name", "Cancel");
        xPSetButton2.setPropertyValue("TabIndex", new Short((short) 0));
        xPSetButton2.setPropertyValue("Label", new String("Cancel"));
        xPSetButton2.setPropertyValue("PushButtonType", new Short((short) PushButtonType.CANCEL_value));

        // create the label model and set the properties
        Object labelModel = xMultiServiceFactory.createInstance(
                "com.sun.star.awt.UnoControlFixedTextModel");
        XPropertySet xPSetLabel = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, labelModel);
        xPSetLabel.setPropertyValue("PositionX", new Integer(5));
        xPSetLabel.setPropertyValue("PositionY", new Integer(5));
        xPSetLabel.setPropertyValue("Width", new Integer(100));
        xPSetLabel.setPropertyValue("Height", new Integer(14));
        xPSetLabel.setPropertyValue("Name", "LABEL");
        xPSetLabel.setPropertyValue("TabIndex", new Short((short) 1));
        xPSetLabel.setPropertyValue("Label", "Select Quality Criteria:");
        // insert the control models into the dialog model
        XNameContainer xNameCont = (XNameContainer) UnoRuntime.queryInterface(
                XNameContainer.class, dialogModel);
        this.xNameCont = xNameCont;
        xNameCont.insertByName("INVIO", buttonModel);
        xNameCont.insertByName("Cancel", buttonModel2);
        xNameCont.insertByName("Label", labelModel);

        Object checkboxModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");
        XPropertySet xpsCHKProperties = createAWTControl(checkboxModel, "Correctness", "Correctness",
                new Rectangle(10, 20, 150, 12));
        xpsCHKProperties.setPropertyValue("TriState", Boolean.FALSE);
        if (qc.isCorrectness()) {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 1));
        } else {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 0));
        }

        checkboxModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");
        xpsCHKProperties = createAWTControl(checkboxModel, "Simplicity", "Simplicity",
                new Rectangle(10, 35, 150, 12));
        xpsCHKProperties.setPropertyValue("TriState", Boolean.FALSE);

        if (qc.isSimplicity()) {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 1));
        } else {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 0));
        }

        checkboxModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");
        xpsCHKProperties = createAWTControl(checkboxModel, "NonAmbiguity", "Non Ambiguity",
                new Rectangle(10, 50, 150, 12));
        xpsCHKProperties.setPropertyValue("TriState", Boolean.FALSE);
        if (qc.isNonAmbiguity()) {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 1));
        } else {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 0));
        }

        checkboxModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");
        xpsCHKProperties = createAWTControl(checkboxModel, "ContentClarity", "Content Clarity",
                new Rectangle(10, 65, 150, 12));
        xpsCHKProperties.setPropertyValue("TriState", Boolean.FALSE);
        if (qc.isContentClarity()) {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 1));
        } else {
            xpsCHKProperties.setPropertyValue("State", new Short((short) 0));
        }
        //xNameCont.insertByName("Completeness", checkboxModel);
        // create the dialog control and set the model
        Object dialog = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialog", m_xContext);
        XControl xControl = (XControl) UnoRuntime.queryInterface(
                XControl.class, dialog);
        XControlModel xControlModel = (XControlModel) UnoRuntime.queryInterface(
                XControlModel.class, dialogModel);
        xControl.setModel(xControlModel);
        // add an action listener to the button control
        XControlContainer xControlCont = (XControlContainer) UnoRuntime.queryInterface(
                XControlContainer.class, dialog);
        Object objectButton = xControlCont.getControl("INVIO");
        XButton xButton = (XButton) UnoRuntime.queryInterface(XButton.class, objectButton);
        xButton.addActionListener(new ActionListenerImpl(xControlCont));

        // create a peer
        Object toolkit = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.Toolkit", m_xContext);
        XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, toolkit);
        XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xControl);
        xWindow.setVisible(false);
        xControl.createPeer(xToolkit, null);
        // execute the dialog
        XDialog xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, dialog);
        xDialog.execute();

        // dispose the dialog
        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, dialog);
        xComponent.dispose();
    }

    private synchronized void createAboutDialog() throws com.sun.star.uno.Exception {
        // get the service manager from the component context
        XMultiComponentFactory xMultiComponentFactory = m_xContext.getServiceManager();
        // create the dialog model and set the properties
        Object dialogModel = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialogModel", m_xContext);
        XPropertySet xPSetDialog = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, dialogModel);
        xPSetDialog.setPropertyValue("PositionX", new Integer(100));
        xPSetDialog.setPropertyValue("PositionY", new Integer(100));
        xPSetDialog.setPropertyValue("Width", new Integer(200));
        xPSetDialog.setPropertyValue("Height", new Integer(100));
        xPSetDialog.setPropertyValue("Title", new String("About CA"));
        // get the service manager from the dialog model
        XMultiServiceFactory xMultiServiceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(
                XMultiServiceFactory.class, dialogModel);

       

        Object buttonModel2 = xMultiServiceFactory.createInstance(
                "com.sun.star.awt.UnoControlButtonModel");
        XPropertySet xPSetButton2 = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, buttonModel2);
        xPSetButton2.setPropertyValue("PositionX", new Integer(80));
        xPSetButton2.setPropertyValue("PositionY", new Integer(80));
        xPSetButton2.setPropertyValue("Width", new Integer(50));
        xPSetButton2.setPropertyValue("Height", new Integer(14));
        xPSetButton2.setPropertyValue("Name", "Cancel");
        xPSetButton2.setPropertyValue("TabIndex", new Short((short) 0));
        xPSetButton2.setPropertyValue("Label", new String("OK"));
        xPSetButton2.setPropertyValue("PushButtonType", new Short((short) PushButtonType.CANCEL_value));

        // create the label model and set the properties
        Object labelModel = xMultiServiceFactory.createInstance(
                "com.sun.star.awt.UnoControlFixedTextModel");
        XPropertySet xPSetLabel = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, labelModel);
        xPSetLabel.setPropertyValue("PositionX", new Integer(5));
        xPSetLabel.setPropertyValue("PositionY", new Integer(5));
        xPSetLabel.setPropertyValue("Width", new Integer(200));
        xPSetLabel.setPropertyValue("Height", new Integer(150));
        xPSetLabel.setPropertyValue("Name", "LABEL");
        xPSetLabel.setPropertyValue("TabIndex", new Short((short) 1));
        String msg = "Prototype of Content Analysis Tool\n\r https://github.com/ISTI-FMT-LearnPAd/ContentAnalysisComponent/\n\r"
                + "Alessio Ferrari - Giorgio O. Spagnolo\n\r FMT ISTI CNR Pisa Italy";
        xPSetLabel.setPropertyValue("Label", msg);
        // insert the control models into the dialog model
        XNameContainer xNameCont = (XNameContainer) UnoRuntime.queryInterface(
                XNameContainer.class, dialogModel);
        this.xNameCont = xNameCont;
        
        xNameCont.insertByName("Cancel", buttonModel2);
        xNameCont.insertByName("Label", labelModel);

        
        // create the dialog control and set the model
        Object dialog = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialog", m_xContext);
        XControl xControl = (XControl) UnoRuntime.queryInterface(
                XControl.class, dialog);
        XControlModel xControlModel = (XControlModel) UnoRuntime.queryInterface(
                XControlModel.class, dialogModel);
        xControl.setModel(xControlModel);
        

        // create a peer
        Object toolkit = xMultiComponentFactory.createInstanceWithContext(
                "com.sun.star.awt.Toolkit", m_xContext);
        XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, toolkit);
        XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xControl);
        xWindow.setVisible(false);
        xControl.createPeer(xToolkit, null);
        // execute the dialog
        XDialog xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, dialog);
        xDialog.execute();

        // dispose the dialog
        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, dialog);
        xComponent.dispose();
    }
    
    private XPropertySet createAWTControl(Object objControl, String ctrlName,
            String ctrlCaption, Rectangle posSize) {

        XPropertySet xpsProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, objControl);
        try {
            xpsProperties.setPropertyValue("PositionX", new Integer(posSize.x));
            xpsProperties.setPropertyValue("PositionY", new Integer(posSize.y));
            xpsProperties.setPropertyValue("Width", new Integer(posSize.width));
            xpsProperties.setPropertyValue("Height", new Integer(posSize.height));
            xpsProperties.setPropertyValue("Name", ctrlName);
            if (ctrlCaption != "") {
                xpsProperties.setPropertyValue("Label", ctrlCaption);
            }

            if ((getNameContainer() != null) && (!getNameContainer().hasByName(ctrlName))) {
                getNameContainer().insertByName(ctrlName, objControl);
            }
        } catch (Exception e) {
            showError(e);
        }
        return xpsProperties;
    }

    public void addStatusListener(com.sun.star.frame.XStatusListener xControl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

    public XNameContainer getNameContainer() {
        return xNameCont;
    }

    public void removeStatusListener(com.sun.star.frame.XStatusListener xControl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
        return m_implementationName;
    }
    
     private void showError(Throwable e) {
    if (testMode) {
      throw new RuntimeException(e);
    }
    String msg = "An error has occurred in CATool "
        + ":\n" + e + "\nStacktrace:\n";
    msg += e.getLocalizedMessage();
    String metaInfo = "OS: " + System.getProperty("os.name") + " on "
        + System.getProperty("os.arch") + ", Java version "
        + System.getProperty("java.version") + " from "
        + System.getProperty("java.vm.vendor");
    msg += metaInfo;
    DialogThread dt = new DialogThread(msg);
   // e.printStackTrace();
    dt.start();
  }

    public boolean supportsService(String sService) {
        int len = m_serviceNames.length;

        for (int i = 0; i < len; i++) {
            if (sService.equals(m_serviceNames[i])) {
                return true;
            }
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.lang.XInitialization:
    public void initialize(Object[] object)
            throws com.sun.star.uno.Exception {
        if (object.length > 0) {
            m_xFrame = (com.sun.star.frame.XFrame) UnoRuntime.queryInterface(
                    com.sun.star.frame.XFrame.class, object[0]);
        }
    }

    private int getId() {
        return 1; //To change body of generated methods, choose Tools | Templates.
    }

    private String getTitle() {
        return "Not supported yet."; //To change body of generated methods, choose Tools | Templates.
    }

    private String getLanguage() {
        return "english"; //To change body of generated methods, choose Tools | Templates.
    }

    /**
     * action listener
     */
    public class ActionListenerImpl implements com.sun.star.awt.XActionListener {

        private int _nCounts = 0;
        private XControlContainer _xControlCont;

        public ActionListenerImpl(XControlContainer xControlCont) {
            _xControlCont = xControlCont;
        }

        // XEventListener
        public void disposing(EventObject eventObject) {
            _xControlCont = null;
        }

        // XActionListener
        public void actionPerformed(ActionEvent actionEvent) {
            // increase click counter
            _nCounts++;

            // set label text
            // Object label = _xControlCont.getControl("Label");
            // XFixedText xLabel = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, label);
            // xLabel.setText("labelprefix" + _nCounts);
            XControl OCorrectness = _xControlCont.getControl("Correctness");
            XControl OSimplicity = _xControlCont.getControl("Simplicity");
            XControl ONonAmbiguity = _xControlCont.getControl("NonAmbiguity");
            XControl OContentClarity = _xControlCont.getControl("ContentClarity");

            XCheckBox XCheckBoxCorrectness = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, OCorrectness);
            XCheckBox XCheckBoxSimplicity = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, OSimplicity);
            XCheckBox XCheckBoxNonAmbiguity = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, ONonAmbiguity);
            XCheckBox XCheckBoxContentClarity = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, OContentClarity);

            if (XCheckBoxCorrectness.getState() > 0) {
                qc.setCorrectness(true);
            } else {
                qc.setCorrectness(false);
            }

            if (XCheckBoxSimplicity.getState() > 0) {
                qc.setSimplicity(true);
            } else {
                qc.setSimplicity(false);
            }

            if (XCheckBoxNonAmbiguity.getState() > 0) {
                qc.setNonAmbiguity(true);
            } else {
                qc.setNonAmbiguity(false);
            }

            if (XCheckBoxContentClarity.getState() > 0) {
                qc.setContentClarity(true);
            } else {
                qc.setContentClarity(false);
            }

            XDialog xDialog = (XDialog) UnoRuntime.queryInterface(
                    XDialog.class, OCorrectness.getContext());

            // Close the dialog
            xDialog.endExecute();
            //xDialog.endExecute();
        }
    }
    
    
     static class DialogThread extends Thread {
    private final String text;

    DialogThread(String text) {
      this.text = text;
    }

    @Override
    public void run() {
        JOptionPane.showMessageDialog(null, text);
    }
  }
    
}

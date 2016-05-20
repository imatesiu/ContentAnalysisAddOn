package cnr.isti;

import cnr.isti.xml.data.QualityCriteria;
import cnr.isti.xml.data.collaborative.AnnotatedCollaborativeContentAnalyses;
import cnr.isti.xml.data.collaborative.CollaborativeContent;
import cnr.isti.xml.data.collaborative.CollaborativeContentAnalysis;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XInitialization;
import com.sun.star.lang.XServiceDisplayName;
import com.sun.star.linguistic2.XLinguServiceEventBroadcaster;
import com.sun.star.linguistic2.XLinguServiceEventListener;
import com.sun.star.linguistic2.XProofreader;
import com.sun.star.task.XJobExecutor;
import java.util.ArrayList;
import java.util.List;
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

    private final List<XLinguServiceEventListener> xEventListeners;

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
        return locales.toArray(new Locale[locales.size()]);
    }

    public boolean hasLocale(Locale arg0) {
        return true;
    }

    public void ignoreRule(String arg0, Locale arg1) {
    }

    public void resetIgnoreRules() {
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

            Client client = ClientBuilder.newClient();
            WebTarget target = client.target("http://contanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent");

            CollaborativeContentAnalysis cca = new CollaborativeContentAnalysis();
            cca.setLanguage(this.getLanguage());

            cca.setCollaborativeContent(new CollaborativeContent(String.valueOf(this.getId()), this.getTitle()));
            cca.getCollaborativeContent().setContentplain(paraText);

            cca.setQualityCriteria(new QualityCriteria());
            cca.getQualityCriteria().setCorrectness(true);
            cca.getQualityCriteria().setSimplicity(true);
            cca.getQualityCriteria().setContentClarity(true);
            cca.getQualityCriteria().setNonAmbiguity(true);
            cca.getQualityCriteria().setCompleteness(false);
            cca.getQualityCriteria().setPresentationClarity(false);

            Entity<CollaborativeContentAnalysis> entity = Entity.entity(cca, MediaType.APPLICATION_XML);
            //GenericEntity<JAXBElement<CollaborativeContentAnalysis>> gw = new GenericEntity<JAXBElement<CollaborativeContentAnalysis>>(cca){};
            Response response = target.request(MediaType.APPLICATION_XML).post(entity);

            String id = response.readEntity(String.class);
            
            client = ClientBuilder.newClient();
		if(id==null){
			id="1";
		}

		target = client.target("http://contanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent/"+id+"/status");
		String 	status ="";
		while (!status.equals("OK")) {


			status = target.request().get(String.class);

			//this.setStatus(status);
			if(status.equals("ERROR"))
				break;
		}
		//log.trace("Status: "+status);

		if(status.equals("OK")){
			
			String ipAddress = null;
			if (ipAddress == null) {
			    
			}
			System.out.println("ipAddress:" + ipAddress);
			
			target = client.target("http://contanalysis.isti.cnr.it:8080").path("lp-content-analysis/learnpad/ca/bridge/validatecollaborativecontent/"+id);
			Response annotatecontent =  target.request().header("X-FORWARDED-FOR", ipAddress).get();
			AnnotatedCollaborativeContentAnalyses res = annotatecontent.readEntity(new GenericType<AnnotatedCollaborativeContentAnalyses>() {});
			//this.setCollectionannotatedcontent(res.getAnnotateCollaborativeContentAnalysis());

		}

        } catch (Throwable t) {

            paRes.nBehindEndOfSentencePosition = paraText.length();
        }

        return paRes;
    }

    public ContentAnalysisAddOn(XComponentContext context) {
        m_xContext = context;
        xEventListeners = new ArrayList<XLinguServiceEventListener>();
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
            if (aURL.Path.compareTo("Command0") == 0) {
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
            if (aURL.Path.compareTo("Command0") == 0) {
                // add your own code here
                return;
            }
        }
    }

    public void addStatusListener(com.sun.star.frame.XStatusListener xControl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

    public void removeStatusListener(com.sun.star.frame.XStatusListener xControl,
            com.sun.star.util.URL aURL) {
        // add your own code here
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
        return m_implementationName;
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

    private Object getId() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private String getTitle() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private String getLanguage() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

}

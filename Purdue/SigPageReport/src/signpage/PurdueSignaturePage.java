package signpage;



import java.io.*;
import java.lang.reflect.*;
import java.net.*;
import java.util.*;

import javax.swing.JOptionPane;

import com.fasttrack.businessobject.*;
import com.fasttrack.dataproxy.*;
import com.fasttrack.debug.*;
//import com.fasttrack.tspd.app.businessproxy.*;
import com.fasttrack.tspd.businessobject.*;
import com.fasttrack.tspd.packager.*;
import com.fasttrack.tspd.proxy.*;
import com.fasttrack.tspd.utilities.*;
import com.fasttrack.tspd.workflow.*;
import com.fasttrack.utilities.*;
import com.sun.xml.internal.bind.v2.model.impl.ArrayInfoImpl;

public class PurdueSignaturePage extends WorkflowAttachment {

	Long _documentPK = new Long(0);
	Long _trialpk;

	public AuditInfo getAuditInfo() {
		return new WorkflowAttachment.AuditInfo(_documentPK, _documentPK, "tspd_document",
				AuditHistCVE.Action.EXPORT, AuditHistCVE.EntityType.DOCUMENT,
				"Signature Page Report");
	}

	public int getInitTimeoutMillis() {
		return WorkflowAttachment.NoTimeout;
	}

	public String getMenuItemName() {

		return "Signature Page Report...";
	}

	public int getRunTimeoutMillis() {
		return WorkflowAttachment.NoTimeout;
	}

	public Response initialize() throws FTException {
		Log.println(PurdueSignaturePage.class, ErrorLevel.INFORMATIONAL,
				"PurdueSignPage, initialize");

		this.addResponseItem(WorkflowAttachment.ExportICP);

		Log.println(PurdueSignaturePage.class, ErrorLevel.INFORMATIONAL,
				"PurdueSignPage, initialize, completed");

		//Getting System TemplatePath
		//
		String strPath = getSystemTemplatePath() + "\\integrations\\commons-lang-2.4.jar";
		try {
			//JOptionPane.showMessageDialog(null, strPath);
			ClassPathHacker.addFile(strPath);
		} catch (Exception e) {
			Log.println(PurdueSignaturePage.class, ErrorLevel.ERROR, "Load class on runtime: "
					+ e.getMessage());
		}

		return new Response(Response.NoError);
	}

	public Response run(WorkflowEvent event, Long pk) throws FTException {

		try {

		
			// Code for exporting Audit Log in to Xml and placing it in to Cache Folder.
			TSPDTrialCO co = queryHelper.getTSPDTrial();
			TSPDDocumentCO doc = co.getTSPDDocument(pk);
			
			String localExportFile = "";
			if (co != null) {
				// local export file
				localExportFile = "";
				_trialpk = queryHelper.getTSPDTrial().getPKValue();
			
				localExportFile = FilePackager.getInstance().getTrialDirectoryPath(co.getTrialId(),
						true).getPath()
						+ "\\" + "AuditExport.xml";
								File outFile = new File(localExportFile);
				
			
				
				AuditHistCVE.Action audEv = null;
				Long userId = null;
				
				
				
				
				
			//	List<DocType> docTypeList = new List<DocType>();
				
				
				List<DocType> docTypeList = new ArrayList<DocType>();

				// you really should get the DocType of this document and add it to the list, so you are only getting audits for that doctype
				//  docTypeList.add(docType)

				DocTypeMgr dtm = DocTypeMgr.getInstance();
				docTypeList.add(dtm.getDocType(doc));
				
				
				CommonObjectVector cov = null;
				
				cov = TSPDProxy.getRunningInstance().loadDocumentAuditLog(new TrialPK(_trialpk),
						new TSPDDocumentPK(pk), new TSPDAuditFilter(0, null, audEv, userId,docTypeList));
				try {
					String xmlOut = SimpleXmlFactory.coToXml(cov);
					writeXml(xmlOut, outFile);
				} catch (Exception ex) {
					String msg = "Signature Page Report, failed exporting Audit log.";
					return new Response(Response.UnknownError, msg);
				}
			}

			if (event == WorkflowAttachment.ExportICP) {
				// Location of executable
				//String installDir = FTBaseZero.getBaseClientPath(ApplicationType.TSPD);

				//String installDir =  FilePackager.getInstance()
				//	.getTrialDirectoryPath(co.getTrialId())
				//	.getPath();

				//Getting current Author Pool
				//DocAuthorCVE authors = TspdDocumentCO.getAuthors();

				String dispName = "";
				String title = "";
				String email = "";
				String mystring = "";

			//	TSPDDocumentCO doc = co.getTSPDDocument(pk);
				CommonObjectVector authors = doc.getAuthors();

			
				
				TSPDProxy proxy = TSPDProxy.getRunningInstance();
				CommonObjectVector users = proxy.loadAuthorList();				
				

				for (Iterator authorIter = authors.iterator(); authorIter.hasNext();) {
					DocAuthorCVE docAuthor = (DocAuthorCVE) authorIter.next();
					//LightWeightUserCVE trialowner = (LightWeightUserCVE) users.findByPrimaryKeyField(docAuthor.getFtUserID());

					Long userID = docAuthor.getFtUserID();
					FtUserCO ftuser = proxy.loadUser(new FtUserPK(userID));
					dispName = ftuser.getDisplayName();
					title = ftuser.getTitle();
					email = ftuser.getEmail();
					mystring += dispName + "~" + title + "~" + email + "^";
				}

				
				// Protocol ID
				String protocolid = co.getProtocolIdentifier();

				String templateDirPath = TSPDUtilities.getSystemTemplatePath(); //EJBUtilities.getCacheFolder(); 	
				//										 
				String command = templateDirPath + "\\integrations\\SignaturePage.exe";

				String filepath = templateDirPath + "\\integrations\\SignaturePage.doc";

				String parameters[] = new String[] { command, protocolid, mystring,
						templateDirPath + "\\integrations", localExportFile };

				//				JOptionPane.showMessageDialog(null, command + " -->" + protocolid + "-->"
				//						+ mystring);

				try {
					int retValue = Runtime.getRuntime().exec(parameters).waitFor();
				} catch (Exception ex) {

					String msg = "PurdueSignPage, run: failed to run the report." + ex.toString();

					return new Response(Response.UnknownError, msg);
				}

			}
		} catch (Exception ex) {
			ex.printStackTrace();
			String msg = "Signature Page, run: " + ex.getMessage() + " Stack Trace :" +  ex.getStackTrace();
			return new Response(Response.UnknownError, msg);
		}

		return new Response(Response.NoError);
	}

	public static String getSystemTemplatePath() {

		String systemTemplatePath = null;

		try {
			// Load up localizations from the system template
			TSPDProxy proxy = TSPDProxy.getRunningInstance();

			if (proxy != null) {
				boolean workingAsTa = proxy.isFtUserTemplateAdminMode()
						|| proxy.isFtUserLibraryAdminMode();

				FilePackager fp = FilePackager.getInstance();
				systemTemplatePath = fp.getTSPDTemplateRootPath(workingAsTa).getPath() + "\\1";
			}
		} catch (Throwable e) {
		}

		if (systemTemplatePath == null) {

			String cacheFolder = EJBUtilities.getCacheFolder();

			// Last case fallback
			File fi1 = new File(cacheFolder);
			File fi2 = new File(fi1, "TSPDTemplates");
			systemTemplatePath = fi2.getPath() + "\\1";
		}

		return systemTemplatePath;
	}

	private void writeXml(String xmlOut, File outFile) throws FTException {
		if (!StringUtils.isBlank(xmlOut)) {

			try {
				// Create file
				PrintWriter pw = new PrintWriter(new FileOutputStream(outFile));
				pw.println(xmlOut);
				pw.close();

			} catch (IOException e) {
				throw new FTException("Exporting Audit error for " + outFile.getAbsolutePath()
						+ ":", e);
			} catch (Exception e) { //Catch exception if any
				throw new FTException("Exporting Audit - file busy or unavailable: "
						+ outFile.getAbsolutePath());
			}
		} else {
			throw new FTException("Exporting Audit: no bytes in audit vector");
		}
	}

	public static class ClassPathHacker {

		private static final Class[] parameters = new Class[] { URL.class };

		public static void addFile(String s) throws IOException {
			File f = new File(s);
			if (!f.exists()) {
				//				JOptionPane.showMessageDialog(null, "File does not exist");
				String msg = "signature page, addToClasspath failed, file not found, for file: " + f
						+ ", cwd: " + f.getCanonicalPath();
				Log.println(PurdueSignaturePage.class, ErrorLevel.INFORMATIONAL, msg);

				return;
			}

			addFile(f);
		}

		//		public static void addFile() throws IOException {
		//			String s = "C:\\fasttrack\\cache\\WorkingTemplates\\1\\integrations"
		//					+ "\\commons-lang-2.4.jar";
		//			File f = new File(s);
		//			addFile(f);
		//		}

		public static void addFile(File f) throws IOException {
			addURL(f.toURI().toURL());
		}

		public static void addURL(URL u) throws IOException {

			URLClassLoader sysloader = (URLClassLoader) ClassLoader.getSystemClassLoader();
			Class sysclass = URLClassLoader.class;

			//	System.out.println(u.getPath());

			try {
				Method method = sysclass.getDeclaredMethod("addURL", parameters);
				method.setAccessible(true);
				method.invoke(sysloader, new Object[] { u });
				//				JOptionPane.showMessageDialog(null, "Successfully loaded");

			} catch (Throwable t) {
				t.printStackTrace();
				throw new IOException("Error, could not add URL to system classloader");
			}

		}
	}

}

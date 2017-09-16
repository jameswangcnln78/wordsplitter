
import java.io.File;
import java.util.List;

import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ch.qos.logback.classic.Level;


public class WordSplitter {

	private static org.slf4j.Logger logger = LoggerFactory.getLogger(WordSplitter.class);	
	static {
		ch.qos.logback.classic.Logger root = (ch.qos.logback.classic.Logger)LoggerFactory.getLogger(ch.qos.logback.classic.Logger.ROOT_LOGGER_NAME);
		root.setLevel(Level.INFO);
	}
	

	
	public static List<Object> findPBetween(MainDocumentPart documentPart,String header) throws Exception{
		
		String xpath = "//w:p[w:r[w:t[contains(text(),'" + header + "')]]]";

		List<Object> first = documentPart.getJAXBNodesViaXPath(xpath, false);
		List<Object> rest = null;
		
		if (first == null || first.size() == 0) {
			logger.warn("Not find any matched P!");
		} else if (first.size() > 1) {
			logger.warn("Not find the only one! Please make sure the condition");
		} else {
			//list.size == 1
			System.out.println("got " + first.size() + " matching " + xpath );
			Object starter = first.get(0);
			Object ender = null;
			
			String condition = ((P)starter).getPPr().getPStyle().getVal();
			logger.debug("The condition is " + condition);
			
			//$ns1[count(.|$ns2) = count($ns2)]
			//TODO:slow https://stackoverflow.com/questions/3835601/how-would-you-find-all-nodes-between-two-h3s-using-xpath
			String ns1Xpath = "//w:p[w:r[w:t[contains(text(),'" + header + "')]]]/following-sibling::node()";
			String ns2Xpath = "//w:p[w:pPr[w:pStyle[@w:val='"+ condition + "']]][2]/preceding-sibling::node()";
			String wholeXpath = ns1Xpath + "[count(.|"+ ns2Xpath + ") = count("+ ns2Xpath+")]";
			rest = documentPart.getJAXBNodesViaXPath(wholeXpath, false);
			if (rest == null || rest.size() == 0) {
				logger.warn("Not find any rest matched P!");
			} else {
				System.out.println("got rest" + rest.size() + " matching " + wholeXpath );
			}

		}
		first.addAll(rest);
		return first;
	}
	
	public static void saveP(List<Object> list,WordprocessingMLPackage srcMdp) throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
		
		for (Object o:list) {
			if (o instanceof P) {
				Object o2 = XmlUtils.unwrap(o);
				String xml = XmlUtils.marshaltoString(o2,true,true);
				//logger.debug("Para xml is " + xml);
				mdp.addParagraph(xml);
			} else {
				logger.warn("Some class is not inserted into new doc" + o.getClass().getName());
			}
		}		
		
		preserveParts(srcMdp,wordMLPackage);
		
		String filename = System.getProperty("user.dir") + "/OUT_hello.docx";
		Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_SAVE_ZIP_FILE); 
		System.out.println("Saved " + filename);
	}
	
	public static void preserveParts(WordprocessingMLPackage srcMdp, WordprocessingMLPackage dstMdp) throws Exception {
		RelationshipsPart rp = srcMdp.getMainDocumentPart().getRelationshipsPart();
		
		
		
		for (Relationship rel: rp.getRelationships().getRelationship()) {
			
			Part p = rp.getPart(rel);
			
			logger.info("Get part :" + p.getPartName().getName());		
		
			// Now try adding it
			dstMdp.getMainDocumentPart().addTargetPart(p, AddPartBehaviour.OVERWRITE_IF_NAME_EXISTS);
		}
	}
	
	public static void main(String[] args) throws Exception{

		String inputfilepath = System.getProperty("user.dir") + "/docs/Docx4j_GettingStarted.docx";
				
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));		
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
				

		saveP(findPBetween(documentPart,"What is docx4j?"),wordMLPackage);
						
	}


}

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;

import org.apache.jmeter.config.Arguments;
import org.apache.jmeter.config.gui.ArgumentsPanel;
import org.apache.jmeter.control.LoopController;
import org.apache.jmeter.control.gui.LoopControlPanel;
import org.apache.jmeter.control.gui.TestPlanGui;
import org.apache.jmeter.engine.StandardJMeterEngine;
import org.apache.jmeter.protocol.http.control.gui.HttpTestSampleGui;
import org.apache.jmeter.protocol.http.sampler.HTTPSamplerProxy;
import org.apache.jmeter.reporters.ResultCollector;
import org.apache.jmeter.reporters.Summariser;
import org.apache.jmeter.save.SaveService;
import org.apache.jmeter.testelement.TestElement;
import org.apache.jmeter.testelement.TestPlan;
import org.apache.jmeter.threads.ThreadGroup;
import org.apache.jmeter.threads.gui.ThreadGroupGui;
import org.apache.jmeter.util.JMeterUtils;
import org.apache.jorphan.collections.HashTree;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestAPICallWithExcelInput {
	
	public static void main (String arv[]){
		
try{
			
			File jmeterHome = new File("C:\\Users\\vedic\\Downloads\\apache-jmeter-3.2\\apache-jmeter-3.2");
			String inputFile = "C:\\Users\\vedic\\Desktop\\workspace\\Zachry\\Zachry-URL's-170527.xlsx";
			String slash = System.getProperty("file.separator");
			Date currDate = new Date();
			if (jmeterHome.exists()) {
	            File jmeterProperties = new File(jmeterHome.getPath() + slash + "bin" + slash + "jmeter.properties");
	            ArrayList<String> urlArr = new ArrayList<String>();
	            ArrayList<String> serviceMethodArr = new ArrayList<String>();
	    		
	            
	            //System.out.println("INPUT FILE="+inputFile);
				
				//workbook instance for XLSX file 
				XSSFWorkbook workbook = new XSSFWorkbook(new File(inputFile));

				//Get correct sheet from the workbook
				XSSFSheet sheet = workbook.getSheet("ZACHRY-SERVICES");
				
				Row row;
				//Cell cell;
				//Iterate through rows. first row is headers, last row is sum
				for (int i=1; i<=sheet.getLastRowNum();i++){
					row = sheet.getRow(i);
					//System.out.println(row.getCell(0)+":"+row.getCell(1));
					urlArr.add(row.getCell(0).getStringCellValue().trim());
					serviceMethodArr.add(row.getCell(1).getStringCellValue().trim());
					System.out.println("URL="+row.getCell(0).getStringCellValue().trim());
					System.out.println("METHOD="+row.getCell(1).getStringCellValue().trim());
				}
	            
	            //JMeter Engine
                StandardJMeterEngine jmeter = new StandardJMeterEngine();
                
                System.out.println(jmeterProperties.getPath());

                //JMeter initialization (properties, log levels, locale, etc)
                JMeterUtils.setJMeterHome(jmeterHome.getPath());
                JMeterUtils.loadJMeterProperties(jmeterProperties.getPath());
              
                JMeterUtils.initLogging();// you can comment this line out to see extra log messages of i.e. DEBUG level
                JMeterUtils.initLocale();
                
                
                // JMeter Test Plan, basically JOrphan HashTree
                HashTree testPlanTree = new HashTree();
                HTTPSamplerProxy examplecomSampler;
                LoopController loopController;
                ThreadGroup threadGroup;
                int numberOfThreads = 1;
                int rampUpTimeInSeconds = 60;
               
                // Test Plan
                TestPlan testPlan = new TestPlan("JMeter Script with Java");
                
                for (int i=0; i<urlArr.size(); i++){
                	
	                examplecomSampler = new HTTPSamplerProxy();
	                examplecomSampler.setDomain("zsapmobdbdev1.zachrycorp.local:");
	                examplecomSampler.setPort(9292);
	                examplecomSampler.setPath(urlArr.get(i));
	                examplecomSampler.setMethod(serviceMethodArr.get(i));
	                examplecomSampler.setName("HTTP Request");
	                examplecomSampler.setProperty(TestElement.TEST_CLASS, HTTPSamplerProxy.class.getName());
	                examplecomSampler.setProperty(TestElement.GUI_CLASS, HttpTestSampleGui.class.getName());
	               
	
	                // Loop Controller
	                loopController = new LoopController();
	                loopController.setLoops(1);
	                loopController.setFirst(true);
	                loopController.setProperty(TestElement.TEST_CLASS, LoopController.class.getName());
	                loopController.setProperty(TestElement.GUI_CLASS, LoopControlPanel.class.getName());
	                loopController.initialize();
	
	                // Thread Group
	                threadGroup = new ThreadGroup();
	                threadGroup.setName("Sample Thread Group");
	                threadGroup.setNumThreads(numberOfThreads);
	                threadGroup.setRampUp(rampUpTimeInSeconds);
	                threadGroup.setSamplerController(loopController);
	                threadGroup.setProperty(TestElement.TEST_CLASS, ThreadGroup.class.getName());
	                threadGroup.setProperty(TestElement.GUI_CLASS, ThreadGroupGui.class.getName());
	
	                testPlan.setSerialized(true);
	                testPlan.setProperty(TestElement.TEST_CLASS, TestPlan.class.getName());
	                testPlan.setProperty(TestElement.GUI_CLASS, TestPlanGui.class.getName());
	                testPlan.setUserDefinedVariables((Arguments) new ArgumentsPanel().createTestElement());
	
	                // Construct Test Plan from previously initialized elements
	                testPlanTree.add(testPlan);
	                HashTree threadGroupHashTree = testPlanTree.add(testPlan, threadGroup);
	                threadGroupHashTree.add(examplecomSampler);
                }
                
                // save generated test plan to JMeter's .jmx file format
                SaveService.saveTree(testPlanTree, new FileOutputStream(jmeterHome + slash + "Zachry.jmx"));

                //add Summarizer output to get test progress in stdout like:
                // summary =      2 in   1.3s =    1.5/s Avg:   631 Min:   290 Max:   973 Err:     0 (0.00%)
                Summariser summer = null;
                String summariserName = JMeterUtils.getPropDefault("summariser.name", "summary");
               
                if (summariserName.length() > 0) {
                    summer = new Summariser(summariserName);
                }
                
                // Store execution results into a .jtl file, we can save file as csv also
                String reportFile = jmeterHome + slash + "Zachry.jtl";
                String csvFile = jmeterHome + slash + "Zachry.csv";

                ResultCollector logger = new ResultCollector(summer);
                logger.setFilename(reportFile);
                ResultCollector csvlogger = new ResultCollector(summer);
                csvlogger.setFilename(csvFile);
                testPlanTree.add(testPlanTree.getArray()[0], logger);
                testPlanTree.add(testPlanTree.getArray()[0], csvlogger);
                // Run Test Plan
                jmeter.configure(testPlanTree);
                jmeter.run();
                
                
                
                System.out.println("Test completed. See " + reportFile + " file for results");
                System.out.println("JMeter .jmx script is available at " + jmeterHome + slash + "Zachry-"+currDate+".jmx");
                
                //TestAPICalls jas = new TestAPICalls();                
                //jas.WriteResultsSummary(jmxFileName,zeroCountURLs,nonZeroCountURLs,totalNumberOfThreads);
                
                System.exit(1);
               }
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	
}

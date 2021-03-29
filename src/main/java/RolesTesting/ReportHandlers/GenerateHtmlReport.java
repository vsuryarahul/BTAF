package ReportHandlers;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

import java.sql.Timestamp;
import java.util.logging.LogManager;

/**
 * This class initializes the objects required to generate the HTML report
 * and allows the user to set the following
 *  Report Location
 *  Report Title
 *  Report Name
 *  Report Theme
 *
 * @author  Sanchitha Kuppusamy
 * @created 11/05/2020
 * @updated 11/10/2020 - Rahul Vanka
 */

public class GenerateHtmlReport {
    private static ExtentSparkReporter htmlReporter;
    private static ExtentReports extent;
    //    public ExtentReports extent;
    //public ExtentTest test;

    public static ExtentReports initializeValues() {
        //TODO - Make path dynamic based on user's windows ID

        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        htmlReporter = new ExtentSparkReporter("C:\\Users\\v-mcoleb\\Desktop\\Reports\\ExtentReport.html");
        extent = new ExtentReports();  //create object of ExtentReports
        extent.attachReporter(htmlReporter);

        //TODO - Make document title dynamic based on user input/ suite name
        //TODO - Make document report name dynamic based on user input/ suite name

        htmlReporter.config().setDocumentTitle("Automation Report"); // Tittle of Report
        htmlReporter.config().setReportName("Role Based Test"); // Name of the report
        htmlReporter.config().setTheme(Theme.DARK);

        return extent;
    }
}

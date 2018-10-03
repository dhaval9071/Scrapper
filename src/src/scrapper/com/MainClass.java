package src.scrapper.com;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class MainClass
{

	public static void main(String[] args) throws Exception
	{

		Document document = Jsoup.connect("https://apps.shopify.com/oberlo/reviews").get();

		// Elements link = document.getElementsByClass("grid__item");
		// Element element = link.get(0);

		XSSFWorkbook workbook = new XSSFWorkbook();
		FileOutputStream out = new FileOutputStream(new File("createworkbook.xlsx"));
		XSSFSheet spreadsheet = workbook.createSheet("Sheet Name");
		XSSFRow row = spreadsheet.createRow(1);

		XSSFCellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor((short) 10);
		Cell cell0 = row.createCell(0);

		cell0.setCellStyle(style);

		cell0.setCellValue("Reviewer Name");
		row.createCell(1).setCellValue("Review Text");
		row.createCell(2).setCellValue("Review Date");

		Elements wholeReview = document.getElementsByClass("review-listing ");
		for (int i = 0; i < 10; i++)
		{
			XSSFRow rowNumber = spreadsheet.createRow(i + 2);
			Elements reviewerNameElements = wholeReview.get(i).getElementsByClass("review-listing-header__text");

			Elements reviewTextElements = wholeReview.get(i).getElementsByClass("review-listing-body__content truncate-content-copy");
			Elements reviewPTag = reviewTextElements.get(0).getElementsByTag("p");

			Elements dateEleElements = wholeReview.get(i).getElementsByClass("review-listing-metadata__item-value");

			Element reviewerName = reviewerNameElements.get(0);
			System.out.println(" Review Name - " + reviewerName.ownText());
			rowNumber.createCell(0).setCellValue(reviewerName.ownText());

			System.out.println(" Review Text - " + reviewPTag.get(0).ownText());
			rowNumber.createCell(1).setCellValue(reviewPTag.get(0).ownText());

			Element dateEle = dateEleElements.get(1);
			System.out.println(" Review Date - " + dateEle.ownText());
			rowNumber.createCell(2).setCellValue(dateEle.ownText());
		}

		workbook.write(out);
		workbook.close();
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}

}

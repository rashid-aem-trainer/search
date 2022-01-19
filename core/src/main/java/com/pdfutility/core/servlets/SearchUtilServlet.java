/*
 *  Copyright 2015 Adobe Systems Incorporated
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */
package com.pdfutility.core.servlets;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.jcr.Node;
import javax.jcr.Session;
import javax.jcr.query.Query;
import javax.servlet.Servlet;
import javax.servlet.ServletException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.apache.sling.api.servlets.SlingSafeMethodsServlet;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.day.cq.search.PredicateGroup;
import com.day.cq.search.QueryBuilder;
import com.day.cq.search.result.SearchResult;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;

/**
 * Servlet that writes some sample content into the response. It is mounted for
 * all resources of a specific Sling resource type. The
 * {@link SlingSafeMethodsServlet} shall be used for HTTP methods that are
 * idempotent. For write operations use the {@link SlingAllMethodsServlet}.
 */
@Component(service = Servlet.class, property = { "sling.servlet.paths=/bin/search/pdfutility",
		"sling.servlet.methods=GET" })
public class SearchUtilServlet extends SlingSafeMethodsServlet {

	private static final long serialVersionUID = -7033019371459280628L;

	ResourceResolver resolver;

	Logger logger = LoggerFactory.getLogger(getClass());

	@Reference
	QueryBuilder queryBuilder;

	Session session;

	private List<String> fulltextList = new ArrayList<>();

	@Override
	protected void doGet(final SlingHttpServletRequest req, final SlingHttpServletResponse resp)
			throws ServletException, IOException {
		try {
			String[] fulltextArr = req.getParameterValues("fulltext");
			if (null != fulltextArr && fulltextArr.length > 0) {
				for (String fullText : fulltextArr) {
					fulltextList.add(fullText.toLowerCase());
				}
			}
			resolver = req.getResourceResolver();
			session = resolver.adaptTo(Session.class);
			Set<String> totalHitSet = new HashSet<>();
			List<String> pdfList = handlePDF();
			List<String> excelList = handleExcel();
			List<String> wordDocList = handleWordDoc();
			List<String> queryResults = getQueryResults();
			totalHitSet.addAll(pdfList);
			totalHitSet.addAll(excelList);
			totalHitSet.addAll(wordDocList);
			totalHitSet.addAll(queryResults);
			resp.setContentType("text/html");
			StringBuilder sb = new StringBuilder();
			for (String hit : totalHitSet) {
				sb.append("<a href=\"" + hit + "\">" + hit + "</a><br>");
			}
			resp.getWriter().write(sb.toString());
		} catch (Exception e) {
			logger.error(e.getMessage());
		} finally {
			resolver.close();
		}
	}

	private List<String> handleExcel() {
		List<String> excelPaths = new ArrayList<>();
		try {
			String query = "SELECT * FROM [dam:Asset] AS s WHERE ISDESCENDANTNODE(s,'/content/dam') AND s.[jcr:content/metadata/dc:format] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'";
			Iterator<Resource> resultItr = resolver.findResources(query, Query.JCR_SQL2);
			while (resultItr.hasNext()) {
				Resource resource = resultItr.next();
				if (null != resource) {
					Boolean isExcelConatinsFullText = false;
					Resource child = resource.getChild("jcr:content/renditions/original/jcr:content");
					Node node = child.adaptTo(Node.class);
					InputStream in = node.getProperty("jcr:data").getBinary().getStream();
					XSSFWorkbook workbook = new XSSFWorkbook(in);
					if (null != workbook) {
						int numberOfSheets = workbook.getNumberOfSheets();
						for (int i = 0; i < numberOfSheets; i++) {
							XSSFSheet sheet = workbook.getSheetAt(i);
							int firstRowNum = sheet.getFirstRowNum();
							int lastRowNum = sheet.getLastRowNum();
							for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
								XSSFRow row = sheet.getRow(rowNum);
								if (null != row) {
									short firstCellNum = row.getFirstCellNum();
									short lastCellNum = row.getLastCellNum();
									for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) {
										XSSFCell cell = row.getCell(cellNum);
										if (null != cell) {
											if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
												String cellValue = cell.getStringCellValue();
												for (String fulltext : fulltextList) {
													if (null != cellValue && (!cellValue.equalsIgnoreCase(""))
															&& cellValue.toLowerCase().equalsIgnoreCase(fulltext)) {
														isExcelConatinsFullText = true;
														break;
													}
												}
											}
										}
										if (isExcelConatinsFullText == true) {
											break;
										}
									}
								}
								if (isExcelConatinsFullText == true) {
									break;
								}
							}
							if (isExcelConatinsFullText == true) {
								break;
							}
						}
					}
					if (isExcelConatinsFullText == true) {
						excelPaths.add(resource.getPath());
					}
				}
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return excelPaths;
	}

	private List<String> handlePDF() {
		List<String> pdfPaths = new ArrayList<>();
		try {
			String query = "SELECT * FROM [dam:Asset] AS s WHERE ISDESCENDANTNODE(s,'/content/dam') AND s.[jcr:content/metadata/dc:format] = 'application/pdf'";
			Iterator<Resource> resultItr = resolver.findResources(query, Query.JCR_SQL2);
			while (resultItr.hasNext()) {
				Resource resource = resultItr.next();
				if (null != resource) {
					Resource child = resource.getChild("jcr:content/renditions/original/jcr:content");
					Boolean isPDFContainsText = false;
					Node node = child.adaptTo(Node.class);
					InputStream in = node.getProperty("jcr:data").getBinary().getStream();
					PdfReader pdfReader = new PdfReader(in);
					int pages = pdfReader.getNumberOfPages();
					for (int i = 1; i <= pages; i++) {
						String pageContent = PdfTextExtractor.getTextFromPage(pdfReader, i);
						for (String fulltext : fulltextList) {
							if (pageContent.toLowerCase().contains(fulltext)) {
								pdfPaths.add(resource.getPath());
								isPDFContainsText = true;
								break;
							}
						}
						if (isPDFContainsText == true) {
							break;
						}
					}
					pdfReader.close();
				}
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return pdfPaths;
	}

	private List<String> handleWordDoc() {
		List<String> wordPaths = new ArrayList<>();
		try {
			String query = "SELECT * FROM [dam:Asset] AS s WHERE ISDESCENDANTNODE(s,'/content/dam') AND s.[jcr:content/metadata/dc:format] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'";
			Iterator<Resource> resultItr = resolver.findResources(query, Query.JCR_SQL2);
			while (resultItr.hasNext()) {
				Resource resource = resultItr.next();
				if (null != resource) {
					Resource child = resource.getChild("jcr:content/renditions/original/jcr:content");
					Boolean isWordContainsText = false;
					Node node = child.adaptTo(Node.class);
					InputStream in = node.getProperty("jcr:data").getBinary().getStream();
					XWPFDocument document = new XWPFDocument(in);
					List<XWPFParagraph> paragraphs = document.getParagraphs();
					for (XWPFParagraph paragraph : paragraphs) {
						String paragraphText = paragraph.getText();
						for (String fulltext : fulltextList) {
							if (paragraphText.toLowerCase().contains(fulltext)) {
								wordPaths.add(resource.getPath());
								isWordContainsText = true;
								break;
							}
						}
						if (isWordContainsText == true) {
							break;
						}
					}
					document.close();
				}
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return wordPaths;
	}

	private List<String> getQueryResults() {
		List<String> queryResult = new ArrayList<>();
		Map<String, Object> predicateMap = new HashMap<>();
		try {
			com.day.cq.search.Query query = queryBuilder.createQuery(PredicateGroup.create(predicateMap), session);
			SearchResult result = query.getResult();
			Iterator<Resource> resItr = result.getResources();
			while (resItr.hasNext()) {
				Resource res = resItr.next();
				queryResult.add(res.getPath());
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return queryResult;
	}

}

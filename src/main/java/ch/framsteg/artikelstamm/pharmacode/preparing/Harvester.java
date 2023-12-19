/**
 *  Copyright 2023 Framsteg GmbH / Olivier Debenath
 *  
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *  
 *   http://www.apache.org/licenses/LICENSE-2.0
 *   
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */
package ch.framsteg.artikelstamm.pharmacode.preparing;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Harvester {

	private static final String PATH_PARAM = "-p";
	private static final String GTIN_PARAM = "-gtin";
	private static final String PHAR_PARAM = "-phar";
	private static final String PATH_MSG = "Enter path to xsl file [{0}]: ";
	private static final String GTIN_MSG = "Enter column number for GTINs [{0}]: ";
	private static final String PHAR_MSG = "Enter column number for PHARs [{0}]: ";
	private static final String PHAR_ERR_MSG = "Missing parameters. The harvester needs three parameters to operate. Aborting...";
	private static final String ENTERED_PATH_MSG = "Path to input file: {0}";
	private static final String ENTERED_GTIN_MSG = "Column number for GTINs: {0}";
	private static final String ENTERED_PHAR_MSG = "Column number for PHARs: {0}";
	private static final String EMPTY_PARAM_ERR_MSG = "Parameter {0} has no value";
	private static final String FINISH_MSG = "Process has successfully finished. {0} lines written to {1}.";
	private static final String FILE_NAME_ROOT = "_pharmacode.csv";
	private static final String DELIMITER = ",";

	private static final Logger logger = LogManager.getLogger(Harvester.class);

	private static HashMap<String, String> values;

	public static void main(String[] args) {
		if (args.length == 6) {
			String path = getParamValue(args, PATH_PARAM);
			int gtinColumnNumber = Integer.parseInt(getParamValue(args, GTIN_PARAM));
			int pharColumnNumber = Integer.parseInt(getParamValue(args, PHAR_PARAM));
			try (Scanner scanner = new Scanner(System.in)) {
				System.out.println(MessageFormat.format(PATH_MSG, path));
				if (scanner.hasNextLine()) {
					String manualInput = scanner.nextLine();
					path = manualInput.equalsIgnoreCase("") ? path : manualInput;
				}
				logger.info(MessageFormat.format(ENTERED_PATH_MSG, path));
				System.out.println(MessageFormat.format(GTIN_MSG, gtinColumnNumber));
				if (scanner.hasNextLine()) {
					String manualInput = scanner.nextLine();
					gtinColumnNumber = manualInput.equalsIgnoreCase("") ? gtinColumnNumber
							: Integer.valueOf(manualInput);
				}
				logger.info(MessageFormat.format(ENTERED_GTIN_MSG, gtinColumnNumber));
				System.out.println(MessageFormat.format(PHAR_MSG, pharColumnNumber));
				if (scanner.hasNextLine()) {
					String manualInput = scanner.nextLine();
					pharColumnNumber = manualInput.equalsIgnoreCase("") ? pharColumnNumber
							: Integer.valueOf(manualInput);
				}
				logger.info(MessageFormat.format(ENTERED_PHAR_MSG, pharColumnNumber));
				readXLSX(path, gtinColumnNumber, pharColumnNumber);
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			logger.info(PHAR_ERR_MSG);
			System.exit(2);
		}
	}

	private static String getParamValue(String[] args, String param) {
		int pos = 0;
		String extractedParamValue = new String();
		for (String s : args) {
			if (param.equalsIgnoreCase(s)) {
				if (args.length < pos + 1) {
					logger.info(MessageFormat.format(EMPTY_PARAM_ERR_MSG, args[pos]));
				} else {
					extractedParamValue = (args[pos + 1]);
				}
			}
			pos++;
		}
		return extractedParamValue;
	}

	private static void readXLSX(String path, int gtinColumnNumber, int pharColumnNumber) throws IOException {

		values = new HashMap<String, String>();

		FileInputStream file = new FileInputStream(new File(path));
		Workbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		int rowCounter = 0;
		for (Row row : sheet) {
			Cell gtin = row.getCell(gtinColumnNumber);
			Cell pharmaCode = row.getCell(pharColumnNumber);
			String gtinValue = new String();
			String pharmaCodeValue = new String();

			if (gtin != null) {
				gtinValue = gtin.toString();
			}
			if (pharmaCode != null && rowCounter > 0) {
				pharmaCodeValue = String.valueOf((int) pharmaCode.getNumericCellValue());
			}
			rowCounter++;
			generatePairs(gtinValue, pharmaCodeValue);

		}
		String filename = writeToFile(path, values);
		workbook.close();
		logger.info(MessageFormat.format(FINISH_MSG, values.size(), filename));
	}

	private static void generatePairs(String gtin, String pharmaCode) {
		if (!gtin.isBlank()) {
			if (StringUtils.isNumeric(pharmaCode)) {
				List<String> gtins = Arrays.asList(gtin.split("  "));
				for (int i = 0; i < gtins.size(); i++) {
					String key = new String();
					String value = new String();
					key = gtins.get(i).replaceAll("\\s", "");
					value = pharmaCode.replaceAll("\\s", "");
					values.put(key, value);
				}
			}
		}
	}

	private static String writeToFile(String path, HashMap<String, String> content) throws IOException {
		Path outputPath = Paths.get(path).getParent();
		String fileName = outputPath.toString() + System.getProperty("file.separator") + System.currentTimeMillis()
				+ FILE_NAME_ROOT;
		File file = new File(fileName);
		BufferedWriter bf = new BufferedWriter(new FileWriter(file));

		for (Map.Entry<String, String> entry : content.entrySet()) {
			bf.write(entry.getKey() + DELIMITER + entry.getValue());
			bf.newLine();
		}
		bf.flush();
		bf.close();
		return fileName;
	}
}

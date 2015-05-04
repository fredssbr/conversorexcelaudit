package br.com.conversor.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class LeitorProperties {

	private Properties prop;

	private final String arquivo = "config.properties";

	private final String DIVISOR = ",";
	private final String KEY_CABECALHO = "cabecalho";

	private String[] cabecalho;

	public LeitorProperties() throws IOException {
		FileInputStream file = new FileInputStream(arquivo);
		prop = new Properties();
		prop.load(file);
		cabecalho = prop.getProperty(KEY_CABECALHO).split(DIVISOR);
	}

	public String[] getCabecalho() {
		return cabecalho;
	}

}

package br.com.conversor.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class LeitorProperties {

	private Properties prop;

	private final String arquivo = "config.properties";

	private final String DIVISOR = ";";
	private final String KEY_CABECALHO = "cabecalho";
	private final String KEY_STATUS_ORIGEM = "statusOrigem";
	private final String KEY_STATUS_DESTINO = "statusDestino";

	private String[] cabecalho;
	private String[] statusOrigem;
	private String[] statusDestino;

	public LeitorProperties() throws IOException {
		FileInputStream file = new FileInputStream(arquivo);
		prop = new Properties();
		prop.load(file);
		cabecalho = prop.getProperty(KEY_CABECALHO).split(DIVISOR);
		statusOrigem = prop.getProperty(KEY_STATUS_ORIGEM).split(DIVISOR);
		statusDestino = prop.getProperty(KEY_STATUS_DESTINO).split(DIVISOR);
	}

	public String[] getCabecalho() {
		return cabecalho;
	}
	public String[] getStatusOrigem() {
		return statusOrigem;
	}
	
	public String[] getStatusDestino() {
		return statusDestino;
	}

}

package br.com.conversor.excel;

import java.awt.BorderLayout;
import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.BorderFactory;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.border.Border;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;

public class Excel {
	
	private String arquivo;	
	private LeitorProperties prop;
	private List<SolicitacaoOrigem> solicitacoesOrigem;
	private List<SolicitacaoDestino> solicitacoesDestino;
	private JProgressBar progressBar;
	private static final String START_STRING = "Status";
	private static final String STRING_ABERTO = "Aberto";
	/**
	 * Construtor que recebe como parâmetro o caminho do arquivo Excel de origem dos dados
	 * @param arquivoExcelOrigem
	 * @throws IOException 
	 */
	public Excel(String arquivoExcelOrigem) throws IOException{
		try{
			this.arquivo = arquivoExcelOrigem;
			this.prop = new LeitorProperties();
			this.solicitacoesDestino = new ArrayList<>();
			this.solicitacoesOrigem = new ArrayList<>();
			this.progressBar = new JProgressBar();
		}catch(IOException e){
			throw new IOException("Não foi possível ler o arquivo de configuração do programa.");
		}
		
	}
	/**
	 * Faz a leitura dos dados do arquivo Excel e os armazena em memória
	 * @param args
	 * @throws IOException
	 */
	public void lerArquivoExcel() throws IOException {
		try{
			FileInputStream file = new FileInputStream(arquivo);
			// Get the workbook instance for XLS file
			HSSFWorkbook workbook = new HSSFWorkbook(file);
	
			// Get first sheet from the workbook
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			// Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheet.iterator();
			
			int max = 0;
			if(sheet.getLastRowNum() > 3){
				max = sheet.getLastRowNum() - 3;
			}
			this.progressBar.setMaximum(max);
			
			
			// Iterate through each rows from first sheet
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// pula as 3 primeiras linhas (nome da planilha e cabeçalhos)
				if (row.getRowNum() < 3) {
					continue;
				}
				
				this.progressBar.setValue(row.getRowNum() - 3);
				// Get iterator to all cells of current row
				Iterator<Cell> cellIterator = row.cellIterator();
				// For each row, iterate through each columns	
				SolicitacaoOrigem solOrigem = new SolicitacaoOrigem();	
				while (cellIterator.hasNext()) {					
					Cell cell = cellIterator.next();
					switch(cell.getColumnIndex()){
					case 1:	
						solOrigem.setChamado(cell.toString());
						break;
					case 2:	
						solOrigem.setDescricao(cell.toString());
						break;
					case 3:	
						solOrigem.setStatus(cell.toString());
						break;
					case 4:	
						solOrigem.setGrupoResponsavel(cell.toString());
						break;
					case 5:	
						solOrigem.setDataHora(cell.getDateCellValue());
						break;
					default:
						break;
					}
				}
				if(solOrigem != null){
					this.solicitacoesOrigem.add(solOrigem);
				}			
			}
			workbook.close();
			file.close();			
			montaListaSolicitacaoDestino();
		}catch(IOException e){
			throw new IOException("Não foi possível ler o arquivo de origem.");
		}catch(IllegalStateException e){
			throw new IOException("Não foi possível ler o arquivo de origem.");
		}
	}
	
	private void montaListaSolicitacaoDestino() throws IOException{
		
		String auxChamado = "";
		for (int i = 0; i < this.solicitacoesOrigem.size(); i++) {
			
			if(this.solicitacoesOrigem.get(i).getChamado() != null && this.solicitacoesOrigem.get(i).getChamado().trim().length() > 0){
				if(!auxChamado.equalsIgnoreCase(this.solicitacoesOrigem.get(i).getChamado())){
					auxChamado = this.solicitacoesOrigem.get(i).getChamado();
					this.solicitacoesOrigem.get(i).setDescricao(STRING_ABERTO);
				}
				
				SolicitacaoDestino solicitacaoDestino = new SolicitacaoDestino();
				String statusAux = "";				
				if(verificaStatusValido(this.solicitacoesOrigem.get(i).getDescricao())){
					solicitacaoDestino.setChamado("NIM110" + String.format("%06d", Integer.parseInt(this.solicitacoesOrigem.get(i).getChamado())));
					solicitacaoDestino.setDataHora(this.solicitacoesOrigem.get(i).getDataHora());
					statusAux = getStatusFromDescription(this.solicitacoesOrigem.get(i).getDescricao());
					solicitacaoDestino.setStatus(statusAux.length() > 0 ? statusAux.concat(" / ").concat(this.solicitacoesOrigem.get(i).getGrupoResponsavel()) : statusAux);				
					this.solicitacoesDestino.add(solicitacaoDestino);				
				}
			}
		}
	}
	
	private String getStatusFromDescription(String pstatus){
		String status = "";
		if(pstatus !=null && pstatus.trim().length() > 0){
			try{
			if(!pstatus.equals(STRING_ABERTO)){
				status = pstatus.substring(pstatus.lastIndexOf("'", pstatus.lastIndexOf("'") -1) + 1, pstatus.lastIndexOf("'"));
			}else{
				status = pstatus;
			}
			}catch(Exception e){
				status = pstatus;
			}
		}
		return getStatusDestinoByStatusOrigem(status);
	}
	
	private boolean verificaStatusValido(String pstatus){
		boolean retorno = false;
		if(pstatus !=null && pstatus.trim().length() > 0){
			retorno = (pstatus.startsWith(START_STRING) && (pstatus.lastIndexOf("'") == pstatus.length()-2)) || pstatus.equalsIgnoreCase(STRING_ABERTO);
		}
		return retorno;
	}
	
	private String getStatusDestinoByStatusOrigem(String pstatus){
		String status = "";
		if(pstatus !=null && pstatus.length() > 0){
			for (int i = 0; i < this.prop.getStatusOrigem().length; i++) {
				if(pstatus.equalsIgnoreCase(this.prop.getStatusOrigem()[i])){
					status = this.prop.getStatusDestino()[i];
					break;
				}
			}
		}
		return status;
	}
	
	private String getStatusDestinoByStatusOrigem(String pstatus){
		String status = "";
		if(pstatus !=null && pstatus.length() > 0){
			for (int i = 0; i < this.prop.getStatusOrigem().length; i++) {
				if(pstatus.equalsIgnoreCase(this.prop.getStatusOrigem()[i])){
					status = this.prop.getStatusDestino()[i];
					break;
				}
			}
		}
		return status;
	}
	
	public List<SolicitacaoOrigem> getSolicitacoesOrigem() {
		return solicitacoesOrigem;
	}
	
	public List<SolicitacaoDestino> getSolicitacoesDestino() {
		return solicitacoesDestino;
	}
	
	public JProgressBar getProgressBar(){
		return progressBar;
	}
	
	public void escreverArquivoExcel(String arquivoExcelSaida) throws IOException{
		try{
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("AUDIT");
			//Create a new row in current sheet
			Row rowCabecalho = sheet.createRow(0);
			
			//Criação da primeira linha - cabeçalho
			for (int i = 0; i < this.prop.getCabecalho().length; i++) {
				//Create a new cell in current row
				Cell cell = rowCabecalho.createCell(i);
				//Set value to new value
				cell.setCellValue(this.prop.getCabecalho()[i]);			
			}
			
			CreationHelper createHelperData = workbook.getCreationHelper();
			CellStyle cellStyleData = workbook.createCellStyle();
			cellStyleData.setDataFormat(createHelperData.createDataFormat().getFormat("dd/MM/yyyy HH:mm:ss"));
			Date dataAux = null;
			this.progressBar.setMaximum(this.solicitacoesDestino.size());
			//Dados
			for (int i = 0; i < this.solicitacoesDestino.size(); i++) {
				Row rowDados = sheet.createRow(i + 1);
				this.progressBar.setValue(i + 1);
				//chamado
				Cell cell0 = rowDados.createCell(0);
				cell0.setCellValue(this.solicitacoesDestino.get(i).getChamado());
				
				//status
				Cell cell1 = rowDados.createCell(1);
				cell1.setCellValue(this.solicitacoesDestino.get(i).getStatus());
				
				//em branco
				Cell cell2 = rowDados.createCell(2);
				cell2.setCellValue("");
				
				
				//marca data e hora
				Cell cell3 = rowDados.createCell(3);
				dataAux = this.solicitacoesDestino.get(i).getDataHora();
				if(dataAux != null){
					cell3.setCellValue(dataAux);
					cell3.setCellStyle(cellStyleData);
				}else{
					cell3.setCellValue("");
				}
	
			}
			//Ajusta o tamanho do texto para a coluna
			for (int i = 0; i < this.solicitacoesDestino.size(); i++) {
				sheet.autoSizeColumn(i);
			}
			
			FileOutputStream out = new FileOutputStream(arquivoExcelSaida);
	        workbook.write(out);
	        workbook.close();
	        out.close();
		}catch(IOException e){
			throw new IOException("Não foi possível gravar o arquivo de gerado.");
		}        
		
	}

	public static void main(String[] args) throws IOException {
	
		JFrame f = new JFrame("Progresso.");
		try{
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setCurrentDirectory(new File("."));
			fileChooser.setAcceptAllFileFilterUsed(false);
			fileChooser.setFileFilter(new FileFilter() {
				public String getDescription() {
					return "Arquivos Excel (.xls)";
				}
				public boolean accept(File f) {
					return f.getName().toLowerCase().endsWith(".xls") || f.isDirectory();
				}
			});
			if (fileChooser.showDialog(null, "Selecione o arquivo para importação") == JFileChooser.APPROVE_OPTION) {
				File fileEntrada = fileChooser.getSelectedFile();
				Excel conversor = new Excel(fileEntrada.getPath());				
				
			    f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			    Container content = f.getContentPane();
			    conversor.getProgressBar().setMinimum(0);
			    conversor.getProgressBar().setValue(0);
			    conversor.getProgressBar().setStringPainted(true);
			    Border borderLeitura = BorderFactory.createTitledBorder("Lendo arquivo...");
			    conversor.getProgressBar().setBorder(borderLeitura);
			    content.add(conversor.getProgressBar(), BorderLayout.NORTH);
			    f.setSize(300, 100);
			    f.setVisible(true);				
				conversor.lerArquivoExcel();
				f.setVisible(false);
				if(fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION){
					Border borderEscrita = BorderFactory.createTitledBorder("Salvando arquivo...");
					conversor.getProgressBar().setBorder(borderEscrita);
					conversor.getProgressBar().setValue(0);
					f.setVisible(true);
					File fileSaida = fileChooser.getSelectedFile();
					conversor.escreverArquivoExcel(fileSaida.getPath());
					f.setVisible(false);					
					JOptionPane.showMessageDialog(null, conversor.getSolicitacoesDestino().size() + " registros exportados.", "Processo finalizado", JOptionPane.INFORMATION_MESSAGE);
					f.dispose();
				}		
				
			}
		}catch(IOException e){
			JOptionPane.showMessageDialog(null, e.getMessage(), "Processo não executado.", JOptionPane.ERROR_MESSAGE);
			f.setVisible(false);
			f.dispose();
		}
		
	}
	
}

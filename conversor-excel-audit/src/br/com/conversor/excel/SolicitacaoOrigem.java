package br.com.conversor.excel;

import java.util.Date;

public class SolicitacaoOrigem implements Comparable<SolicitacaoOrigem>{

	private String chamado;
	private String descricao;
	private String status;
	private String grupoResponsavel;
	private Date dataHora;

	public String getDescricao() {
		return descricao;
	}

	public void setDescricao(String descricao) {
		this.descricao = descricao;
	}

	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getGrupoResponsavel() {
		return grupoResponsavel;
	}

	public void setGrupoResponsavel(String grupoResponsavel) {
		this.grupoResponsavel = grupoResponsavel;
	}

	public Date getDataHora() {
		return dataHora;
	}

	public void setDataHora(Date dataHora) {
		this.dataHora = dataHora;
	}

	public String getChamado() {
		return chamado;
	}

	public void setChamado(String chamado) {
		this.chamado = chamado;
	}

	@Override
	public int compareTo(SolicitacaoOrigem o) {
		try {
			Integer chamadoThis = new Integer(this.chamado);
			Integer chamadoParam = new Integer(o.getChamado());
			int compareChamado = chamadoThis.compareTo(chamadoParam);
			int compareData = this.dataHora.compareTo(o.getDataHora());
			if(compareChamado==0){
				return compareData;
			}else{
				return compareChamado;
			}
		} catch (NumberFormatException e) {
			return 1;
		}catch (NullPointerException e) {
			return 1;
		}
		
	}

}

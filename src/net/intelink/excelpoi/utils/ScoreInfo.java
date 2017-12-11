package net.intelink.excelpoi.utils;

public class ScoreInfo {
	private String name;
	private String banji;
	private Integer cgrade;
	private Integer wgrade;

	public ScoreInfo() {
	}

	public ScoreInfo(String name, String banji, Integer cgrade, Integer wgrade) {
		super();
		this.name = name;
		this.banji = banji;
		this.cgrade = cgrade;
		this.wgrade = wgrade;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getBanji() {
		return banji;
	}

	public void setBanji(String banji) {
		this.banji = banji;
	}

	public Integer getCgrade() {
		return cgrade;
	}

	public void setCgrade(Integer cgrade) {
		this.cgrade = cgrade;
	}

	public Integer getWgrade() {
		return wgrade;
	}

	public void setWgrade(Integer wgrade) {
		this.wgrade = wgrade;
	}

	@Override
	public String toString() {
		return "ScoreInfo [name=" + name + ", banji=" + banji + ", cgrade=" + cgrade + ", wgrade=" + wgrade + "]";
	}

}

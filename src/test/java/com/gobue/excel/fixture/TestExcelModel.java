package com.gobue.excel.fixture;

import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @author hurong
 * 
 * @description: Add the class description here.
 */
public class TestExcelModel {

    private Long sku;
    private String productName;
    private Integer dc;
    private Boolean isRep;
    private Double cr;
    private Character band;
    private Double formulaRst;
    private Date createDate;
    // 动态列:Excel某列开始，数据形式都一样，用此数据Map存储
    private Map<Double, Double> dcQty;
    private List<Integer> dcQty4Export;

    public Long getSku() {

        return sku;
    }

    public void setSku(Long sku) {

        this.sku = sku;
    }

    public String getProductName() {

        return productName;
    }

    public void setProductName(String productName) {

        this.productName = productName;
    }

    public Integer getDc() {

        return dc;
    }

    public void setDc(Integer dc) {

        this.dc = dc;
    }

    public Boolean getIsRep() {

        return isRep;
    }

    public void setIsRep(Boolean isRep) {

        this.isRep = isRep;
    }

    public Double getCr() {

        return cr;
    }

    public void setCr(Double cr) {

        this.cr = cr;
    }

    public Character getBand() {

        return band;
    }

    public void setBand(Character band) {

        this.band = band;
    }

    public Date getCreateDate() {

        return createDate;
    }

    public void setCreateDate(Date createDate) {

        this.createDate = createDate;
    }

    public Double getFormulaRst() {

        return formulaRst;
    }

    public void setFormulaRst(Double formulaRst) {

        this.formulaRst = formulaRst;
    }

    public Map<Double, Double> getDcQty() {

        return dcQty;
    }

    public void setDcQty(Map<Double, Double> dcQty) {

        this.dcQty = dcQty;
    }

    public List<Integer> getDcQty4Export() {

        return dcQty4Export;
    }

    public void setDcQty4Export(List<Integer> dcQty4Export) {

        this.dcQty4Export = dcQty4Export;
    }

}

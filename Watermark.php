<?php
namespace watermark;

use ZipArchive;


/**
 * excel添加水印，适用于office2007以后的基于openXml的excel即xlsx格式
 * 
 * @version 1.0
 * 
 * @example
 *          $water  = new \watermark\Watermark('D:/a.xlsx');
 *          $num = $water->addImage('D:/images/b.png');
 *          $water->getSheet(1)->setBgImg($num);
 *          $water->close();
 * 
 * @author lxj <my2233@foxmail.com>
 */
class Watermark {
    
    /**
     * @var ZipArchive对象
     */
    private $zip;
    
    /**
     * @var 图像序号
     */
    private $num = 1;
    
    /**
     * @var sheet序号
     */
    private $sheet = 1;
    
    /**
     * @var 图片后缀数组
     */
    private $suffixArr = [];
    
    /**
     * @var 图片命名
     */
    private $nameArr = [];


    /**
     * 初始化
     * 
     * @param string $file 压缩包文件名
     * 
     * @access public
     */
    public function __construct($file = null) {
        if (!empty($file)) {
            $this->zip  = new ZipArchive();
            $this->zip->open($file);
        }
    }
    
    /**
     * 读取zip文件
     * 
     * @param string $file 压缩包文件名
     * 
     * @access public
     */
    public function setFile($file) {
        $this->zip  = new ZipArchive();
        return $this->zip->open($file);
    }
    
    /**
     * 添加图片
     * 
     * @param string $file 压缩包文件名
     * @return int 添加进来的图像的编号
     * 
     * @access public
     */
    public function addImage($file) {
        $suffix = pathinfo($file, PATHINFO_EXTENSION);
        $name   = uniqid();
        $this->zip->addFile($file, 'xl/media/bgimage' . $name . '.' . $suffix);
        $num    = $this->num;
        $this->suffixArr[$num]  = $suffix;
        $this->nameArr[$num]    = $name;
        $this->num++;
        return $num;
    }
    
    /**
     * 获取当前sheet
     * 
     * @param int $num sheet编号
     * @return object
     * 
     * @access public
     */
    public function getSheet($num = 1) {
        $this->sheet = $num;
        return $this;
    }
    
    /**
     * 设置背景图
     * 
     * @param int $num 图像编号
     * 
     * @access public
     */
    public function setBgImg($num) {
        $nowSuffix  = strtolower($this->suffixArr[$num]);
        $this->zip->addFromString("xl/worksheets/_rels/sheet$this->sheet.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bgimage'.$this->nameArr[$num].'.'.$nowSuffix.'"/></Relationships>');
        $string = $this->zip->getFromName("xl/worksheets/sheet$this->sheet.xml");
        $string = str_replace('</worksheet>', '<picture r:id="rId1"/></worksheet>', $string);
        $this->zip->addFromString("xl/worksheets/sheet$this->sheet.xml", $string);
        
        $str1   = $this->zip->getFromName('[Content_Types].xml');
        if (!strpos($str1, 'Extension="' . $nowSuffix . '"') || !strpos($str1, 'ContentType="image/' . $nowSuffix . '"')) {
            $str1   = str_replace('</Types>', '<Default ContentType="image/'. $nowSuffix .'" Extension="' . $nowSuffix . '"/></Types>', $str1);
            $this->zip->addFromString("[Content_Types].xml", $str1);
        }
    }
    
    /**
     * 关闭
     * 
     * @access public
     */
    public function close() {
        $this->zip->close();
    }
    
}

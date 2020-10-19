# excel-watermark
add watermark to excel 给excel添加背景图片水印


### Install via composer
$water  = new \watermark\Watermark('D:/a.xlsx');
$num = $water->addImage('D:/images/b.png');
$water->getSheet(1)->setBgImg($num);
$water->close();

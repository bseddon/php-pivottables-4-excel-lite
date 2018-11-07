<?php
function load_data() {
  $xmlDoc = new \DOMDocument();
  $xmlDoc->load( __DIR__ . "/data.xml");
  $data = array();
  foreach ($xmlDoc->documentElement->childNodes AS $item)
  {
    if ( $item->nodeType == XML_ELEMENT_NODE ) {
    $data []= array( "Account" => $item->getAttribute("account"),
      "Genre" => $item->getAttribute("genre"),
      "Images" => $item->getAttribute("images"),
      "Average Ranking" => $item->getAttribute("avgrank"),
      "Total Size" => $item->getAttribute("size") );
    }
  }
  return $data;
}
?>
<?php
/**
 * class cointains methods that return values to replace tokens with.
 */
class companyTokens{
	static public function name() {
		return 'Company Name';
	}
	static public function address1() {
		return "123 Example St.";
	}
	static public function city() {
		return "Winnipeg";
	}
	static public function prov() {
		return "MB";
	}
	static public function postalcode() {
		return 'R3P 2G1';
	}
	static public function phone() {
		return '123-555-1234';
	}
}
/**
 * class cointains methods that return values to replace tokens with.
 */
class clientTokens{
	static public function company() {
		return 'Client X';
	} 
	static public function name() {
		return "Mr. X Smith";
	}
	static public function city() {
		return 'Winnipeg';
	}
	static public function prov() {
		return "MB";
	}
	static public function postalcode() {
		return "R3M 1Z4";
	}
}
/**
 * class cointains methods that return values to replace tokens with.
 */
class productTokens {
	static public function desc() {
		return array( 'Product A', 'Product B' );
	}
	static public function rate() {
		return array( '30.63', '30.63' );
	}
	static public function qty() {
		return array( '1', '2' );
	}
	static public function total() {
		return array( "30.63", ''.round( 2 * 30.63, 2 ) );
	}
}
ini_set( "display_errors", 1 );
ini_set( "error_reporting", E_ALL );
//create a soap client that connects to the Python Mail Merge Service 
$pyMailMergeService = new SoapClient( null, array( 'uri'=>'urn:approve','location'=>"http://localhost:8888/" ) );
//request a list of tokens that are available in the given odt
$tokens = $pyMailMergeService->getTokens( 'invoice.odt' );
$params = array();
foreach( $tokens as $token ) {
	//split the token on :: to find what part should be the class and what part should be the method
	list( $tmp, $method ) = explode( "::", $token );
	//check for modifiers, and remove them.  ie. the repeat row, modifier. PHP does not need to worry about this.  Only python.
	$modifier = '';
	if( strpos( $tmp, '|' ) === false )
		$class = $tmp;
	else
		list( $modifier, $class ) = explode( "|", $tmp );
	//build the classname.
	$class .= "Tokens";
	//call the method, if it exists.
	if( class_exists( $class ) )
		if( method_exists( $class, $method ) )
			$params[$token] = call_user_method_array( $method, $class, array() );
}
//set the fileoutput to pdf, the service should be able to genderate anything that OpenOffice.org can handle, I've tried, doc, rtf and html
$doctype = 'pdf'; 
//generate the document.
$doc = $pyMailMergeService->convert( 'invoice.odt', $params, $doctype );
//write the file to the disk.
//the doc comes back base64 encoded, because I was getting errors sending binary data via soap. Something about utf-8 chars only.
file_put_contents( '/var/www/html/invoice.'.$doctype, base64_decode( $doc ) );
?>
<?php
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
$python = new SoapClient( null, array( 'uri'=>'urn:approve','location'=>"http://localhost:8888/" ) );
$tokens = $python->getTokens( 'invoice.odt' );
$app = new stdClass();
$params = array();
foreach( $tokens as $token ) {
	list( $tmp, $method ) = explode( "::", $token );
	$modifier = '';
	if( strpos( $tmp, '|' ) === false )
		$class = $tmp;
	else
		list( $modifier, $class ) = explode( "|", $tmp );
	$class .= "Tokens";
	if( method_exists( $class, $method ) )
		$params[$token] = call_user_method_array( $method, $class, array( $app, $modifier ) );
}
$pdf = $python->pdf( 'invoice.odt', $params );
file_put_contents( '/var/www/html/invoice'.".pdf", base64_decode( $pdf ) );
?>
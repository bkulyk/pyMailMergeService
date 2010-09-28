<?php
/**
 * Connect to pyMailMergeService in order to create PDF (or other) documents via OpenOffice.org
 * @package MailMergeService
 */
class MailMergeService{
	/**
	 * @var SoapClient $connection //hold the global connection object, so there is only one ever created at a time
	 */
	static protected $connection = null;
	/**
	 * @var String $uri //uri information necessary for the SOAP connection
	 */
	static protected $uri = "urn:approve";
	/**
	 * @var String $host //host name for the SOAP server.
	 */
	static protected $host = "http://localhost:8888/";
	/**
	 * Get the connection object for the mail merger service
	 * @return SoapClient
	 * @since May 2010
	 */
	protected function getConnection() {
		if( is_null( self::$connection ) )
			return self::connect();
		else
			return self::$connection;
	}
	/**
	 * Connect to the Mail Merge Service via SOAP.
	 * @return SoapClient
	 * @since May 2010
	 */
	static private function connect() {
		self::$connection = new SoapClient( null, array( 'uri'=>self::$uri,'location'=>self::$host ) );
		return self::$connection;
	}
	/**
	 * Ask the mail merge service to open an odt file and determine what tokens are available 
	 * to use in it.
	 * @param String $odtFilename //name of open office document on MailMergeService server to get the tokens from
	 * @return String[] //example array( 'invoice::date', 'products::name|1' )
	 * @since May 2010
	 */
	public function getTokens( $odtFilename ) {
		$result = self::getConnection()->getTokens( $odtFilename );
		if( !is_array( $result ) ) {
			if( substr( $result, 0, 6 ) == 'error:' ) {
				throw new MMSException( $result );
				return false;
			}
		}else
			return $result;
	}
	/**
	 * Ask the mail merge service to merge the data from $params into the openoffice document ($odtFilename)
	 * and then convert it into the $doctype docment format.
	 * @param String $odtFilename //name of open office document on MailMergeService server to merge the data into
	 * @param Assoc $params //Associate array where the keys are the tokens from MailMergeService::getTokens and the value is a string or array of strings or a 1 or 0
	 * @param String $doctype [optional] //examples: odt, pdf, doc
	 * @return binary //Converted document as a raw string.
	 * @since May 2010
	 */
	public function convert( $odtFilename, $params, $doctype='pdf' ) {
		$result = self::getConnection()->convert( $odtFilename, $params, $doctype );
		if( substr( $result, 0, 6 ) == 'error:' ) {
			throw new MMSException( $result );
			return false;
		}else
			return base64_decode( $result );
	}
	/**
	 * Same as convert only the type is contrained to PDF
	 * @param String $odtFilename //name of open office document on MailMergeService server to merge the data into
	 * @param Assoc $params //Associate array where the keys are the tokens from MailMergeService::getTokens and the value is a string or array of strings or a 1 or 0
	 * @return binary //Converted pdf document as a raw string.
	 * @since May 2010
	 */
	public function pdf( $odtFilename, $params ) {
		$result = self::getConnection()->pdf( $odtFilename, $params );
		if( substr( $result, 0, 6 ) == 'error:' ) {
			throw new MMSException( $result );
			return false;
		}else
			return base64_decode( $result );
	}
	public function upload( $odtFilePath ) {
		if( file_exists( $odtFilePath ) ) {
			$noContent = false;
			$zip = zip_open( $odtFilePath );
			if( is_resource( $zip ) ) {
				$notZip = false;
				//make sure that it's a valid open office document.
				do{
	                $entry = zip_read( $zip );
	            }while( $entry && zip_entry_name( $entry ) != "content.xml" );
				if( zip_entry_name( $entry ) == "content.xml" ) {
					//do upload
				}else
					$noContent = true;
			}else
				$notZip = true;
			if( $notZip || $noContent ) 
				throw new Exception( "File provided is not a valid ODT: $odtFilePath" );
		}else
			throw new Exception( "ODT file does not exist: $odtFilePath" );
	}
}
class MMSException extends Exception{ }
?>
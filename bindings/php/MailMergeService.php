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
	static public $host = "http://localhost:8888/";
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
			return new MailMergeDocument( $result, $doctype, true );
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
			return new MailMergeDocument( $result, 'pdf', true );
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
				throw new MMSUploadException( "File provided is not a valid ODT: $odtFilePath" );
		}else
			throw new MMSUploadException( "ODT file does not exist: $odtFilePath" );
	}
}
/**
 * Wrapper for the document that's returned from the MailMergeService. 
 * I decided to include this class because it was a real pain trying to determine the correct 
 * headers to output a file to the browser and still have it work with all browsers
 * (specificially IE6 w/ SSL)
 * @package MailMergeService
 */
class MailMergeDocument {
	protected $document = null;
	protected $docType = null;
	//@todo add more document types.  This is just the bare minimum that I need
	static protected $docTypeMap = array( "pdf"=>"application/pdf", "odt"=>"application/vnd.oasis.opendocument.text" );
	/**
	 * The document comes back from the mail merge service as base64 encoded, because soap doesn't like
	 * to send content that is not utf-8.  So decode it.
	 * @param String $documentContent
	 * @param String $doctype
	 * @param Boolean $base64ed
	 */
	public function __construct( $documentContent, $doctype, $base64ed=true ) {
		$this->docType = $doctype;
		if( $base64ed )
			$this->document = base64_decode( $documentContent );
		else
			$this->document = $documentContent;
	}
	/**
	 * This is some backword compatability for the company I work for.  I had already 
	 * written all the code to use the document as a string, so this will prevent that code 
	 * from breaking.
	 */
	public function __toString() {
		return $this->document;
	}
	/**
	 * Output the file to the browser as an attachment.
	 * @param String $filename //name of the file as it would appear on the user's pc.
	 * @param Boolean $sendHeaders //true to send the headers for the user to download the document
	 * as a file attachment.  The only real reason you wouldn't want to do this is if, you needed 
	 * some custom headers, or were already sending them.
	 */
	public function output( $filename=null, $sendHeaders=true ) {
		if( !empty( $this->document ) ) {
			if( !empty( $filename ) && $sendHeaders ) {
				$this->sendHeaders( $filename );
			}
		}
		echo $this->document;
	}
	/**
	 * Output the headers, to he browser.  These are specifiially tested so they work in all browsers.
	 * Getting ie6 (with SSL) was a real pain in the butt.
	 * @param String $filename
	 * @throws MMSDocumentException
	 */
	public function sendHeaders( $filename ) {
		if( empty( $this->document ) )
			throw new MMSDocumentException( "Cannot send headers for a document with no content." );
		header( "Content-type: ".$this->getMime() );
		header( "Pragma: Public", true );
		header( "Expires: ".gmdate( "D, d M Y H:i:s", time() )." GMT", true );
		header( "Content-disposition: attachment; filename=$filename" );
		if( strlen( $this->document ) )
			header( "Content-Length: ".strlen( $this->document ), true);
	}
	/**
	 * Get the mime type for the file extension.
	 * @throws MMSDocumentException
	 * @return String
	 */
	public function getMime() {
		if( empty( $this->docType ) )
			throw new MMSDocumentException( "Document has no docType.  Cannot determine mime type" );
		return self::$docTypeMap[ $this->docType ];
	}
}
class MMSException extends Exception {}
class MMSDocumentException extends Exception {}
class MMSUploadException extends Exception {}
?>

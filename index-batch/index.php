<?php
error_reporting(E_ALL);
require '../vendor/autoload.php';

use League\OAuth2\Client\Provider\GenericProvider;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model\MailFolder;
use Microsoft\Graph\Model\Message;


// Get an access token from Microsoft Graph, simple enough with League's OAuth 2 Client
$accessToken = static function (){

    $oauthClient = new GenericProvider([
        'clientId'                  => '',
        'clientSecret'              => '',
        'redirectUri'               => 'http://localhost:8000/callback',
        'urlAuthorize'              => 'https://login.microsoftonline.com/6e26a76f-adf1-4e9e-8eea-25f26bf52a0f/oauth2/v2.0/authorize',
        'urlAccessToken'            => 'https://login.microsoftonline.com/6e26a76f-adf1-4e9e-8eea-25f26bf52a0f/oauth2/v2.0/token',
        'urlResourceOwnerDetails'   => '',
        'scope'                    => 'https://graph.microsoft.com/.default'
    ]);

    return $oauthClient->getAccessToken('client_credentials', ['scope' => 'https://graph.microsoft.com/.default'])->getToken();

};

// Instantiate a new instance of the Microsoft Graph SDK and set the access token.
$graph = new Graph();
$graph->setAccessToken($accessToken());

// Query the archive folder in my exchange online mailbox.
// This returns an object containing information such as
// the number of items in the folder.
$archive = $graph->createRequest(
    'GET',
    "/users/cb922f30-bd4d-48f9-b3e3-c9b892c6294e/mailFolders('Archive')"
)->setReturnType(MailFolder::class)->execute();

// Initialise a new variable with an array to store any folder names with their corresponding Ids.
// This is to reduce the amount of times I have to call Microsoft Graph.
$stored_folder_ids = [];

// If there are no child folders and there are no messages in the archive folder, there is nothing to process.
if ($archive->getChildFolderCount() === 0 && $archive->getTotalItemCount() === 0){
    return 'Nothing to Process';
}

// Get all the child folders of the archive folder.
// I would normally have put some pagination logic here
// but since there are two folders and the default
// page size is 10 i'm good for testing purposes.
$childFolders = $graph->createCollectionRequest('GET', "/users/cb922f30-bd4d-48f9-b3e3-c9b892c6294e/mailFolders/Archive/childFolders")
    ->setReturnType(MailFolder::class)
    ->getPage();

// Store the folder display name as the key and
// make the value the id for each childFolder
// into the stored_folder_ids array.
array_walk($childFolders, static function(MailFolder $folder) use (&$stored_folder_ids){
    $stored_folder_ids[$folder->getDisplayName()] = $folder->getId();
});

// This is were the problems begin.
// Here I instantiate a collection request from Microsoft Graph,
// I only execute it once i'm inside the while loop.
$messageGrabber = $graph->createCollectionRequest('GET', "/users/cb922f30-bd4d-48f9-b3e3-c9b892c6294e/mailFolders('Archive')/messages")
    ->setReturnType(Message::class)
    ->setPageSize(4);

$iterationCount = 1;

while(!$messageGrabber->isEnd()){

    echo "<h1>Page: {$iterationCount}</h1>";

    $messages = $messageGrabber->getPage();

    $batchRequestBody = [
        'requests' => []
    ];

    $batchRequestCounter = 0;

    foreach ($messages as $message){

        $batchRequestBody['requests'][] = [
            'id' => $batchRequestCounter++,
            'method' => 'POST',
            'url' => "/users/cb922f30-bd4d-48f9-b3e3-c9b892c6294e/messages/{$message->getId()}/move",
            'body' => [
                'destinationId' => 'AAMkADZhYjI0ZTMwLTIzNTMtNDgzMS04ODJjLTNhMjAzYzYwY2NlNgAuAAAAAAAENC3I2XUATZdPFyH_DYPLAQC-CUleI1JNQpuxqWV3wPhOAAAmEWoDAAA='
            ],
            'headers' => [
                'Content-Type' => 'application/json'
            ]
        ];

        echo "<p>Message: {$message->getId()}</p>";
    }

    $iterationCount++;

    $graph->createRequest('POST', '/$batch')
        ->attachBody($batchRequestBody)
        ->execute();

    $responses[] = $batchRequestBody;

}
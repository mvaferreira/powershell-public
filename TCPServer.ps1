$port = 7569;
$Listener = [System.Net.Sockets.TcpListener]$port;
$Listener.Start();

while($true)
{
	$client = $Listener.AcceptTcpClient();
	Write-Host "Connection on port $port from $($client.client.RemoteEndPoint.Address)";
	$client.Close();
}
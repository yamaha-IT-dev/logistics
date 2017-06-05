<%
sub ListFolderContents(path)
	dim fs, folder, file, item, url

    set fs = CreateObject("Scripting.FileSystemObject")
    set folder = fs.GetFolder(path)

    'Display the target folder and info.

    Response.Write("Summary: " _
       & folder.Files.Count & " files, ")
    if folder.SubFolders.Count > 0 then
       Response.Write(folder.SubFolders.Count & " directories, ")
    end if
    Response.Write(Round(folder.Size / 1024) & " KB total." _
       & vbCrLf)

    Response.Write("<ul>" & vbCrLf)

    'Display a list of sub folders.

    for each item in folder.SubFolders
    	ListFolderContents(item.Path)
    next

    'Display a list of files.

    for each item in folder.Files
       	url = MapURL(item.path)
    	Response.Write("<li><a href=""" & url & """>" & item.Name & "</a>" & vbCrLf)
    next

	Response.Write("</ul>" & vbCrLf)
    Response.Write("</li>" & vbCrLf)
end sub

function MapURL(path)
    dim rootPath, url

    'Convert a physical file path to a URL for hypertext links.

    rootPath = Server.MapPath("/")
    url = Right(path, Len(path) - Len(rootPath))
    MapURL = Replace(url, "\", "/")
end function

%>
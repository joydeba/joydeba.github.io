function MakePreviewImagePopup( imgName, fileName, contextId, link, item, ou){

	var popup = new D2L.Dialog('PreviewImage');
	var opener;
	if( item != null ){
		opener = item.GetStructure().GetOpener();
	}
	popup.SetTitle(new D2L.LP.Text.LangTerm("eP_Common.Shared.lblPreviewArtifactDialogSubject", imgName));
	popup.SetSize('375','380');
	
	
	var imgContainer = new D2L.Control.Container()
	var img = popup.CreateElement('img');
	img.src = "/d2l/eP/artifacts/context_artifact_stream.d2lFile?ou=" + ou + "&contextId=" + contextId + "&width=350&height=250";
	imgContainer.AppendChild( img );
	popup.AppendChild( imgContainer );
	
	var navInfo = new D2L.NavInfo();
	navInfo.SetNavigation( link );
	
	var linkContainer = new D2L.Control.Container();
	
	var linkControl = new D2L.Control.Link()
	linkControl.SetText( new D2L.Language.Term( "eP_Common.Shared.altDownloadSubject", "" ) );
	linkControl.SetAlt( new D2L.Language.Term( "eP_Common.Shared.altDownloadSubject", imgName ) );
	linkControl.SetNav(navInfo);
	
	var img2Control = new D2L.Control.Image();
	img2Control.SetImage( new D2L.Images.ImageTerm( "eP_Common.Main.actDownload" ));
	img2Control.SetAlt( new D2L.Language.Term( "eP_Common.Shared.altDownloadSubject", imgName ) );
	img2Control.AppendTo( linkContainer );
	linkControl.AppendTo( linkContainer );
	popup.AppendChild(linkContainer);
	popup.SetHasCloseIcon( true );
	popup.AddButton( D2L.Control.Button.Type.Close);
	popup.Open( opener );
}

function ViewArtifact(link, contextChain) {
    var rpc = D2L.Rpc.Create('MarkArtifactViewed', undefined, '/d2l/eP/dashboard/rpc/rpc_functions.d2l');
    rpc.Call(contextChain);
    var ni = new D2L.NavInfo();
    ni.SetNavigation(link);
    ni.SetTarget('_blank');
    Nav.Go(ni);
}

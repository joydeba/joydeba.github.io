D2L.eP.Common.Sharing.Quickshare = function( objectId, sharableObjects ){

	var rpcCallback = function( rpcResponse ){

	}
	 var rpc = D2L.Rpc.Create( 'CreateShares', rpcCallback,'/d2l/ep/ePObjects/sharing/rpc/rpc_functions.d2l' ).Call( objectId, sharableObjects );
};

var PresentationPageInfo = D2L.Class.extend({
		Construct: function(){
			arguments.callee.$.Construct.call( this );
			 this.PageId = 0;
			 this.PageName = '';
			 this.PageLayoutId = 0;
			 },
			 Deserialize: function(deserializer) {
			 this.PageId = deserializer.GetMember( 'PageId' );
			 this.PageName = deserializer.GetMember( 'PageName' );
			 this.PageLayoutId = deserializer.GetMember( 'PageLayoutTypeId' );
			 }
});

var ClientSideEPObject = D2L.Class.extend({
		Construct: function(){
			arguments.callee.$.Construct.call( this );

			 this.ObjectId = 0;
			 this.ObjectType = 0;
			 this.CreateDate = null;
			 this.LastModifiedDate = null;
			 this.Name = '';
			 this.Description ='';
			 this.UserName = '';
			 this.UserId = 0;
			 this.OrgId = 0;
		},

		Deserialize: function(deserializer) {
			this.ObjectId = deserializer.GetMember( 'ObjectId' );
			 this.ObjectType = deserializer.GetMember( 'ObjectType' );
			 this.CreateDate = deserializer.GetMember( 'CreateDate' );
			 this.LastModifiedDate = deserializer.GetMember( 'LastModfiedDate' );
			 this.Name = deserializer.GetMember( 'Name' );
			 this.Description = deserializer.GetMember( 'Description' );
			 this.UserName = deserializer.GetMember( 'UserName' );
			 this.UserId = deserializer.GetMember( 'UserId' );
			 this.OrgId = deserializer.GetMember( 'OrgId' );
		}
});

D2L.eP.Common.HandleAddTo = function( draggedObject, targetObject ){
	//Handle Add To Presentation
	if( draggedObject.GetData("DDType") == "IEPObjectEntity"){
		if( targetObject.GetData("DDType") == "IEPObjectEntity"){
			switch( targetObject.GetData("Type") ){
				case D2L.eP.Common.EPObjectType.Presentation:
					AddToPresentationDialog (targetObject.GetData("Id"), draggedObject.GetData("Id"), draggedObject.GetData("Type") );
				break;
				case D2L.eP.Common.EPObjectType.Collection:
					new D2L.Rpc.Create( 'AddToCollection', null ,'/d2l/ep/common/controls/EPObjectEntityDataList.control.d2l').Call( targetObject.GetData("Id"), draggedObject.GetData("Id"));
				break;
					case D2L.eP.Common.EPObjectType.Reflection:
				break;
			}
		}else if( targetObject.GetData("DDType") == "ISharable"){
			D2L.eP.Common.Sharing.HandleDropEvent( targetObject, draggedObject );
		}
	}else if( draggedObject.GetData("DDType") == "ISharable" &&  targetObject.GetData("DDType") == "IEPObjectEntity" ){
		D2L.eP.Common.Sharing.HandleDropEvent( draggedObject, targetObject  );
	}
}

D2L.eP.Common.Sharing.ISharableSelectedContainer = function( sharableItem, node, showImage ){

	if( showImage ){
		var imageContainer = new D2L.Control.Container();
		imageContainer.SetFloat( D2L.Style.Float.Left )
		node.AppendChild(  imageContainer );
		var image = new D2L.Control.Image();
		if( sharableItem.ProfileType === D2L.eP.Common.Sharing.PermissionProfileType.Individual){
			image.SetImage( new D2L.Images.ImageTerm( 'CMC.Main.tbManageCourses' ));
		}else if( sharableItem.ProfileType === D2L.eP.Common.Sharing.PermissionProfileType.OrgUnit){
			image.SetImage( new D2L.Images.ImageTerm( 'ManageUsers.Main.tbUsers' ));
		}else{
			image.SetImage( new D2L.Images.ImageTerm( 'CMC.Main.actGroups' ));
		}
		imageContainer.AppendChild( image );
	}else{
		node.SetFloat( D2L.Style.Float.Left );
	}
	var txtContainer = new D2L.Control.Container();
	txtContainer.SetFloat( D2L.Style.Float.Left );
	var txtName = new D2L.Control.TextBlock( new D2L.LP.Text.PlainText( sharableItem.DisplayName ));
	txtContainer.AppendChild( txtName );
	node.AppendChild( txtContainer );

}

D2L.eP.Common.Sharing.ISharableItemRendererContainer = function( sharableItem, showImage, contextId, layoutType, onUpdateCallback ){

	var node = new D2L.Control.Container();
	var txtName = new D2L.Control.TextBlock(new D2L.LP.Text.SmlText(sharableItem.Name));
	//showimage apparently determines which list this item is in, whether it is the candidates, or the selected list.
	if (showImage){
		txtName = new D2L.Control.TextBlock(new D2L.LP.Text.SmlText(sharableItem.DisplayName));
	}

	if( layoutType != null && layoutType == D2L.Control.DataList.LayoutTypes.Tiles){
		node.SetSize( 80, 80 );
		node.SetTextAlignment( D2L.Style.TextAlignment.Justified );
		if( showImage ){
			var imageContainer = new D2L.Control.Container();
			imageContainer.SetFloat( D2L.Style.Float.None );
			node.AppendChild(  imageContainer );
			var clf = new D2L.Control.ClearFloat();
			node.AppendChild( clf );
			if( sharableItem.ProfileType === D2L.eP.Common.Sharing.PermissionProfileType.Individual){
				var upb = new D2L.Control.UserProfileBadge( sharableItem.GetUserProfileInfo()) ;
				imageContainer.SetTextAlignment( D2L.Style.TextAlignment.Middle );
				upb.AppendTo( imageContainer);
			} else {
				var topPadding = new D2L.Control.Container();
				topPadding.SetHeight( 15 );
				imageContainer.AppendChild( topPadding );
				imageContainer.SetTextAlignment( D2L.Style.TextAlignment.Middle );
				var image = new D2L.Control.Image();
				image.SetImage( sharableItem.GetImage() );
				image.AppendTo( imageContainer);
			}
		}
	} else {
	if( showImage ){
		var imageContainer = new D2L.Control.Container();
		imageContainer.SetFloat( D2L.Style.Float.Left )
		imageContainer.SetSize( 40,40 );
		node.AppendChild(  imageContainer );

		if( sharableItem.ProfileType === D2L.eP.Common.Sharing.PermissionProfileType.Individual){
				var upb = new D2L.Control.UserProfileBadge( sharableItem.GetUserProfileInfo()) ;

				if( upb.m_userProfileBadgeInfo.GetProfileBadgeImage().m_size != undefined){
					upb.m_userProfileBadgeInfo.GetProfileBadgeImage().m_size = 35;
				}else{
					upb.m_userProfileBadgeInfo.GetProfileBadgeImage().m_width = 35;
					upb.m_userProfileBadgeInfo.GetProfileBadgeImage().m_height = 35;
				}
				imageContainer.SetTextAlignment( D2L.Style.TextAlignment.Middle );
				upb.AppendTo( imageContainer);
		} else {
				var image = new D2L.Control.Image();
				image.SetImage( sharableItem.GetImage() );
				image.m_height = 35;
				image.m_width = 35;
				imageContainer.SetTextAlignment( D2L.Style.TextAlignment.Middle );
				imageContainer.AppendChild( image );
			}
		}
	}
	var txtContainer = new D2L.Control.Container();
	txtContainer.SetFloat( D2L.Style.Float.Left );

	txtContainer.AppendChild( txtName );
	txtContainer.AppendChild( new D2L.Control.ClearFloat() );
	if( sharableItem.Special && sharableItem.Name != sharableItem.Category ){
		//Add email address under External Users name
		txtContainer.AppendChild(
			D2L.eP.Common.Sharing.CreateExtraInfo( sharableItem.Category )
		);
	} else if( sharableItem.ExtraInfo != false ) {
		// Display extra info
		txtContainer.AppendChild(
			D2L.eP.Common.Sharing.CreateExtraInfo( sharableItem.ExtraInfo )
		);
	}

	// Show the Reset Invite when has objectId, is expired, and is external user (special + profileType)
	if ( contextId && sharableItem.Expired
			&& sharableItem.Special && sharableItem.ProfileType == 3 ) {
		txtContainer.AppendChild( new D2L.Control.ClearFloat() );
		txtContainer.AppendChild(
			D2L.eP.Common.Sharing.CreateResetExternalInvite( sharableItem, contextId, onUpdateCallback )
		);
	}

	node.AppendChild( txtContainer );
	node.AppendChild( new D2L.Control.ClearFloat() );



	return node;
}

D2L.eP.Common.Sharing.CreateExtraInfo = function( text ) {
	var extraInfoContainer = new D2L.Control.Container();
	extraInfoContainer.SetFloat( D2L.Style.Float.Left );
	var extraInfoText = new D2L.Control.TextBlock( new D2L.LP.Text.SmlText( text ));
	extraInfoContainer.AppendChild( extraInfoText );
	return extraInfoContainer;
}

D2L.eP.Common.Sharing.CreateResetExternalInvite = function( shareableItem, contextId, onUpdateCallback ) {
	var resetContainer = new D2L.Control.Container();
	resetContainer.SetFloat( D2L.Style.Float.Left );

	var text = new D2L.Control.TextBlock(
			new D2L.LP.Text.LangTerm( "eP_PermissionProfiles.Shared.lblExternalExpired" )
		);
	
	var link = new D2L.Control.Link();
	link.SetText( new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.lblResetAvailability' ) );
	link.SetOnClick( function() {
		ResetExternalInvite( shareableItem, contextId, onUpdateCallback );
	} );

	resetContainer.AppendChild( text );
	resetContainer.AppendChild( link );

	return resetContainer;
}

D2L.eP.Common.Sharing.AddToSelectedList = function( iSharable, selectedIds, onCompleteCallback){
	if(!(selectedIds.ContainsKey( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType))){
		 selectedIds.Add( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType, iSharable );
	 }
	if(onCompleteCallback){
		onCompleteCallback();
	}
}

D2L.eP.Common.Sharing.RemoveFromSelectedList = function( iSharable, selectedIds, onCompleteCallback ){
	if(selectedIds.ContainsKey( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType)){
		 selectedIds.Remove( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType);
	 }
	if(onCompleteCallback){
		onCompleteCallback();
	}
}

D2L.eP.Common.Sharing.AddToUpdatedList = function( iSharable, updatedIds ){
	updatedIds.Add( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType, iSharable );
}
D2L.eP.Common.Sharing.RemoveFromUpdatedList = function( iSharable, updatedIds ){
	updatedIds.Remove( iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType );
}


D2L.eP.Common.Sharing.ConstructSelectedListControl = function( selectedIds, listControl, onCompleteCallback){
  var quickShareList = listControl;
	quickShareList.RemoveAll();
	for( s in selectedIds.Values() ){
		 var iSharable = selectedIds.Values()[s];
		 var newNode = new D2L.Control.ListItem( );
		 var outerContainer = new D2L.Control.Container();

		 var clf = new D2L.Control.ClearFloat();

		 var container =  D2L.eP.Common.Sharing.ISharableItemRendererContainer( iSharable, false );
		 container.SetFloat( D2L.Style.Float.Left );

		 outerContainer.SetFloat( D2L.Style.Float.Left );
		 outerContainer.GetPadding().SetAllValues( 1 );
		 outerContainer.GetSpacing().GetLeft().SetValue( 4 );
		 outerContainer.GetSpacing().GetRight().SetValue( 4 );
		 outerContainer.SetBackgroundColour('#e8e8ff');

		 var border = outerContainer.GetBorder();

		 border.SetStyle( D2L.Style.BorderStyle.Dotted );
		 border.SetColour( '#dedede' );

		 border.SetWidth( 1 );

		 outerContainer.SetBorder( border );

		 var id = iSharable.Name + '_' + iSharable.Id +'_'+ iSharable.ProfileType;
		 var deleteImage = new D2L.Control.Image();
		 deleteImage.SetImage( new D2L.Images.ImageTerm( 'Shared.Main.actRemove' ));
		 var x = D2L.eP.Common.Sharing.GetCurrentSharable( iSharable, selectedIds, function(){ onCompleteCallback(); } );
		 deleteImage.SetOnClick( x );

		 var imageContainer = new D2L.Control.Container();
		 imageContainer.SetFloat( D2L.Style.Float.Right );
		 imageContainer.GetSpacing().GetLeft().SetValue( 3 );
		 imageContainer.AppendChild( deleteImage );
		 imageContainer.AppendChild( clf )

		 outerContainer.AppendChild( container );
		 outerContainer.AppendChild( imageContainer );
		 outerContainer.AppendChild( clf );

		 newNode.AppendChild( outerContainer );
		 quickShareList.AddItem( newNode );
	}
}

D2L.eP.Common.Sharing.GetCurrentSharable = function( iSharable, selectedIds, onCompleteCallback ){
	var x = function(){
					D2L.eP.Common.Sharing.RemoveFromSelectedList( iSharable, selectedIds ,onCompleteCallback) ;
			 }
	return x;
}



D2L.eP.Common.Sharing.IPermission = D2L.Class.extend({
		Construct: function(){
			arguments.callee.$.Construct.call( this );
			this.Init();
		},

		Init: function(){
			 this.m_permissions = new D2L.Util.Dictionary();
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.View , true );
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.SeeComments ,false );
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.AddComments ,false );
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.ViewAssessments ,false );
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.AddAssessments ,false );
			 this.m_permissions.Add( D2L.eP.Common.Sharing.ObjectRightType.Edit ,false );
		},

		Serialize: function(serializer){
			serializer.AddMember( 'Permissions', this.m_permissions );
		},

		SetHasPermission: function( objectRightType, value ){
			this.m_permissions.Add(objectRightType, value );
			this.m_permissionHaveChanged = true;
		},

		GetHasPermission: function( objectRightType ){
			return this.m_permissions.Get( objectRightType );
		},

		AddPermission: function( objectRightType ){
			this.SetHasPermission( objectRightType, true );
		},

		RemovePermission: function( objectRightType ){
			this.SetHasPermission( objectRightType, false );

		},
		GetPermissionsHaveChanged: function(){
			return this.m_permissionHaveChanged;
		},
		GetPermissionsArray: function(){
			var results = [];
			var i = 0;
			for( s in this.m_permissions.Keys()){
				if(this.m_permissions.Get(this.m_permissions.Keys()[s])){
					results[i] = this.m_permissions.Keys()[s];
					i++;
				}
			}
			return results;
		}
});

D2L.eP.Common.Sharing.ISharable = D2L.Class.extend({
	Construct: function() {
		arguments.callee.$.Construct.call(this);
		this.Id = 0;
		this.TargetId = 0;
		this.Name = '';
		this.DisplayName = '';
		this.Category = '';
		this.Special = false;
		this.Deletable = true;
		this.ExpirationDate = null;
		this.Expired = false;
		this.ProfileType = null;
		this.ProfileContext = null;
		this.Permissions = new D2L.eP.Common.Sharing.IPermission();
		this.ProfileInfo = null;
		this.Image = null;
		this.ExtraInfo = '';
	},

	Deserialize: function(deserializer) {
		this.Id = deserializer.GetMember('Id');
		this.TargetId = deserializer.GetMember('TargetId');
		this.Name = deserializer.GetMember('Name');
		this.DisplayName = deserializer.GetMember('DisplayName');
		this.Category = deserializer.GetMember('Category');
		this.Special = deserializer.GetMember('Special');
		this.Deletable = deserializer.GetMember('Deletable');
		this.Expired = deserializer.GetMember('Expired');
		this.ProfileType = deserializer.GetMember('ProfileType');
		this.ProfileContext = deserializer.GetMember('ProfileContext');
		
		if (deserializer.HasMember('ExpirationDate')) {
			this.ExpirationDate = new Date( deserializer.GetMember('ExpirationDate') );
		}
		if (deserializer.HasMember('ProfileInfo')) {
			this.ProfileInfo = deserializer.GetObject('ProfileInfo'); //, D2L.Control.UserProfileBadgeInfo );
		}
		if (deserializer.HasMember('Image')) {
			this.Image = deserializer.GetObject('Image', D2L.Images.Image);
		}
		if (deserializer.HasMember('ExtraInfo')) {
			this.ExtraInfo = deserializer.GetMember('ExtraInfo');
		}

		var permissions = deserializer.GetObjectArray('Permissions');
		if (permissions != null) {
			this.Permissions = new D2L.eP.Common.Sharing.IPermission();
			this.Permissions.Init();
			for (p in permissions) {
				this.Permissions.AddPermission(permissions[p]);
			}
		}


	},

	Serialize: function(serializer) {

		serializer.AddMember('Id', this.Id);
		serializer.AddMember('TargetId', this.TargetId);
		serializer.AddMember('Name', this.Name);
		serializer.AddMember('DisplayName', this.DisplayName);
		serializer.AddMember('Category', this.Category);
		serializer.AddMember('Deletable', this.Deletable);
		serializer.AddMember('Expired', this.Expired);
		serializer.AddMember('Special', this.Special);
		serializer.AddMember('ProfileType', this.ProfileType);
		serializer.AddMember('ProfileContext', this.ProfileContext);
		serializer.AddMember('Permissions', this.Permissions.GetPermissionsArray());

		if ( this.ExpirationDate ) {
			serializer.AddMember( 'ExpirationDate', this.ExpirationDate.toJSON() );
		}
	},
	GetImage: function() {
		return this.Image;
	},
	GetUserProfileInfo: function() {
		return this.ProfileInfo;
	},
	GetPermissions: function() {
		return this.Permissions;
	}

});





D2L.eP.Common.Sharing.CreateSharingPermissions = function( defaultSharableObject, selectedIds, profileType, onAddCallback ){
	var x = function(){

		var resp =function( rpcResponse ){
			var results = [];
			if ( rpcResponse.GetResponseType() == D2L.Rpc.ResponseType.Success ) {
				  var results =  rpcResponse.GetResult();
				}
			if (onAddCallback) {
				onAddCallback( results, selectedIds, defaultSharableObject.Id );
			}
		}

		 var sharables = [];
		 var i = 0;

		 for( s in selectedIds.Values() ){
			 sharables[i] = selectedIds.Values()[s];
			 sharables[i].Permissions = defaultSharableObject.Permissions;
			 i++;
		 }

		 new D2L.Rpc.Create('CreateShares', resp, '/d2l/ep/ePObjects/sharing/rpc/rpc_functions.d2l').Call(defaultSharableObject.Id, sharables, profileType);
	}
	return x;
}

D2L.eP.Common.Sharing.UpdatePermission = function(saveControl, control, permissionsObject) {
    var x = function() {
        var objectType = control.GetValue();
        var value = control.IsChecked();
        permissionsObject.SetHasPermission(objectType, value);
        D2L.eP.Common.Sharing.UpdatePermissionSaveState(saveControl);
    };
    return x;
};


D2L.eP.Common.Sharing.UpdatePermissionSaveState = function(saveControl) {
    if (saveControl) {
        saveControl.SetIsEnabled(true);
    }
};

D2L.eP.Common.Sharing.UpdateSharingPermissions = function(iSharableArray, callback) {
    var x = function() {
        new D2L.Rpc.Create('UpdateShares', callback, '/d2l/ep/ePObjects/sharing/rpc/rpc_functions.d2l').Call(iSharableArray);
    };
    return x;
};

D2L.eP.Common.Sharing.DeleteShares = function(iSharableArray, contextId, callback) {
    var x = function() {
        new D2L.Rpc.Create('DeleteShares', callback, '/d2l/ep/ePObjects/sharing/rpc/rpc_functions.d2l').Call(iSharableArray, contextId);
    };
    return x;
};


function GenerateRowPermissionsBox(controlId, sharableObject, onSaveCallback) {
	if ( onSaveCallback == null ) {
		return D2L.eP.Common.Sharing.PermissionTextBlockControl( sharableObject );
	}

	var container = new D2L.Control.Container();
	var saveLinkContainer = new D2L.Control.Container();

	container.SetFloat( D2L.Style.Float.Right );
	saveLinkContainer.SetFloat( D2L.Style.Float.Left );

	var a = [];
	a[0] = sharableObject;

	var saveLink = new D2L.Control.Link();
	saveLink.SetFloat(D2L.Style.Float.Right);
	saveLink.SetText(new D2L.LP.Text.LangTerm('Standard.Buttons.btnSave'));

	var usp = D2L.eP.Common.Sharing.UpdateSharingPermissions( a, onSaveCallback );
	saveLink.SetOnClick( usp );
	saveLinkContainer.AppendChild( saveLink );
	container.AppendChild( saveLinkContainer );
	var result = GeneratePermissionsBox( controlId, sharableObject, saveLink, true );
	result.Children()[1].AppendChild( container );
	return result;
}

function GenerateNewPermissionSetBox( controlId, defaultSharableObject, selectedIds, profileType, onAddCallback, addFunctionOut, attachToDialog ){
	var showButton = !addFunctionOut;

	var a = [];
	a[0] = defaultSharableObject;
	var usp = D2L.eP.Common.Sharing.CreateSharingPermissions(defaultSharableObject, selectedIds, profileType, onAddCallback);

	if (addFunctionOut) {
		addFunctionOut.val = usp; // Make the 'add' function available to the caller.
	}

	var result = GeneratePermissionsBox(controlId, defaultSharableObject, addButton, false);

	if (showButton) {
		var addButtonContainer = new D2L.Control.Container();

		var addButton = new D2L.Control.Button();
		addButton.SetText(new D2L.LP.Text.LangTerm('Standard.Buttons.btnAdd'));
		addButton.SetIsPrimary(true);
		addButton.SetStyle(D2L.Control.Button.Style.Action);
		addButton.SetOnClick(usp);

		addButtonContainer.AppendChild(addButton);
		result.AppendChild(addButtonContainer);
	}

	return result;
}

function GeneratePermissionsBox( controlId, sharableObject, saveLink, useEditButton ){
	var outerContainer = new D2L.Control.Container();
	outerContainer.SetFloat( D2L.Style.Float.Left );

	var clf = new D2L.Control.ClearFloat();
	var displayContainer = new D2L.Control.Container();
	var container = new D2L.Control.Container();
	var txtTerms = new D2L.Util.Dictionary();



	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.View , [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optView' ),  new D2L.Images.ImageTerm('eP_Common.Main.actAssessments')]);
	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.SeeComments, [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optSeeComments' ),  new D2L.Images.ImageTerm('eP_Common.Main.actComments')]);
	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.AddComments, [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optAddComments' ),  new D2L.Images.ImageTerm('eP_Common.Main.infHasComments')]);
	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.ViewAssessments, [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optSeeAssessments' ),  new D2L.Images.ImageTerm('eP_Common.Main.actAssessments')]);
	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.AddAssessments, [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optAddAssessments' ),  new D2L.Images.ImageTerm('eP_Common.Main.infHasAssessments')]);
	txtTerms.Add( D2L.eP.Common.Sharing.ObjectRightType.Edit, [new D2L.LP.Text.LangTerm( 'eP_PermissionProfiles.Shared.optEdit' ),  new D2L.Images.ImageTerm('eP_Common.Main.actAssessments')]);

	outerContainer.AppendChild(displayContainer);
	outerContainer.AppendChild(container);

	if (useEditButton) {
		displayContainer.SetIsDisplayed( true );
		container.SetIsDisplayed( false );

		var editImageContainer = new D2L.Control.Container();
		editImageContainer.SetFloat(D2L.Style.Float.Right);

		var editImage = new D2L.Control.Image();
		editImage.SetImage(new D2L.Images.ImageTerm('Shared.Main.actEdit'));
		editImage.SetAlt(new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.altEditSharePermission', sharableObject.Name));

		var toggleVis = ToggleVisibility(container, displayContainer);
		editImage.SetOnClick(toggleVis);

		editImageContainer.AppendChild(editImage);
	} else {
		displayContainer.SetIsDisplayed( false );
		container.SetIsDisplayed( true );
	}

	displayContainer.AppendChild( D2L.eP.Common.Sharing.PermissionTextBlockControl(sharableObject) );
	displayContainer.AppendChild( editImageContainer );

	var oType = D2L.eP.Common.Sharing.ObjectRightType.SeeComments;
	var canViewCommentsContainer = D2L.eP.Common.Sharing.PermissionCheckboxControl( controlId + "_" + oType, sharableObject, oType, txtTerms.Get(oType)[0], true, saveLink );
	oType = D2L.eP.Common.Sharing.ObjectRightType.AddComments;
	var canAddCommentsContainer = D2L.eP.Common.Sharing.PermissionCheckboxControl( controlId + "_" + oType, sharableObject, oType, txtTerms.Get(oType)[0], true, saveLink );
	oType = D2L.eP.Common.Sharing.ObjectRightType.ViewAssessments;
	var canViewAssessmentsContainer = D2L.eP.Common.Sharing.PermissionCheckboxControl( controlId + "_" + oType, sharableObject, oType, txtTerms.Get(oType)[0], true, saveLink );
	oType = D2L.eP.Common.Sharing.ObjectRightType.AddAssessments;
	var canAddAssessmentsContainer = D2L.eP.Common.Sharing.PermissionCheckboxControl( controlId + "_" + oType, sharableObject, oType, txtTerms.Get(oType)[0], true, saveLink );

	oType = D2L.eP.Common.Sharing.ObjectRightType.Edit;
	// External Users can't receive edit permissions
	var editIsEnabled = !(sharableObject.Special && sharableObject.ProfileType === D2L.eP.Common.Sharing.PermissionProfileType.Individual);
	var canEditContainer = D2L.eP.Common.Sharing.PermissionCheckboxControl( controlId + "_" + oType, sharableObject, oType, txtTerms.Get(oType)[0], editIsEnabled, saveLink );

	container.AppendChild( canViewCommentsContainer );
	container.AppendChild( clf );
	container.AppendChild( canAddCommentsContainer );
	container.AppendChild( clf );
	container.AppendChild( canViewAssessmentsContainer );
	container.AppendChild( clf );
	container.AppendChild( canAddAssessmentsContainer );
	container.AppendChild( clf );
	container.AppendChild( canEditContainer );
	container.AppendChild( clf );

	return outerContainer;
}

D2L.eP.Common.Sharing.PermissionCheckboxControl = function( controlId, sharableObject, objectRightType, controlText, isEnabled, saveLink ){
	var cbControlContainer = new D2L.Control.Container();
	var cbControl = new D2L.Control.Checkbox();
	cbControl.SetControlId( controlId );
	cbControl.SetValue( objectRightType );
	cbControl.SetIsChecked( sharableObject.GetPermissions().GetHasPermission( objectRightType ) );
	cbControl.SetText( controlText  );
	cbControl.SetIsEnabled( isEnabled );
	var cv = D2L.eP.Common.Sharing.UpdatePermission( saveLink, cbControl, sharableObject.GetPermissions() );
	cbControl.SetOnClick( cv );
	cbControlContainer.AppendChild( cbControl );
	return cbControlContainer;
};


D2L.eP.Common.Sharing.PermissionTextBlockControl = function (sharableObject, bold) {
    var result = new D2L.Control.Container();
    result.SetFloat(D2L.Style.Float.Left);
    var spacer;

    new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.lblPermissionTypeSpacer').GetText().Register(function (text) {
        spacer = text;
    });


    if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.View)) {
        new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionViewExpanded').GetText().Register(function (text) {
            result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(text)));
        });
    }
    if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.AddComments) && sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.SeeComments)) {
        new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionAddSeeCommentsExpanded').GetText().Register(function (text) {
            result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
        });
    } else {
        if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.AddComments)) {
            new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionAddCommentsExpanded').GetText().Register(function (text) {
                result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
            });
        }
        if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.SeeComments)) {
            new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionSeeCommentsExpanded').GetText().Register(function (text) {
                result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
            });
        }
    }
    if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.AddAssessments) && sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.ViewAssessments)) {
        new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionAddSeeAssessmentsExpanded').GetText().Register(function (text) {
            result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
        });
    } else {
        if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.AddAssessments)) {
            new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionAddAssessmentsExpanded').GetText().Register(function (text) {
                result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
            });
        }
        if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.ViewAssessments)) {
            new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.hdrPermissionSeeAssessmentsExpanded').GetText().Register(function (text) {
                result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
            });
        }
    }
    if (sharableObject.GetPermissions().GetHasPermission(D2L.eP.Common.Sharing.ObjectRightType.Edit)) {
        new D2L.LP.Text.LangTerm('eP_PermissionProfiles.Shared.optEdit').GetText().Register(function (text) {
            result.AppendChild(new D2L.Control.TextBlock(new D2L.LP.Text.PlainText(spacer + text)));
        });
    }

    return result;
};


ToggleVisibility = function (enableControl, disableControl) {
    var x = function () {
        enableControl.SetIsDisplayed(true);
        disableControl.SetIsDisplayed(false);
    };
    return x;
};

function MakeObjectPermissionPopup( title, ou, contextId, opener ){
	var dialog = new D2L.Dialog( );
	dialog.SetTitle(new D2L.LP.Text.LangTerm('eP_PermissionProfiles.ObjectSharing.titItemNameSharingSettings', title));
	dialog.SetSize('700', '700');
	dialog.SetSrc('/d2l/ep/epobjects/sharing/sharing.d2l', 'SrcCallback');
	dialog.SetHasCloseIcon(true);

	dialog.SetCallback(
		function (dialogResponse) {
			dialogResponse.GetDialog().Close();
		});

	dialog.AddCustomButton(
			new D2L.LP.Text.LangTerm('Standard.Buttons.btnClose'),
			D2L.Dialog.ButtonPosition.Left,
			D2L.Dialog.ResponseType.Positive,
			'1',
			'done',
			true
		);
	dialog.SetSrcParam('contextId', contextId);
	dialog.SetSrcParam('objectTitle', title);
	dialog.SetSrcParam('ou', ou);
	dialog.Open(opener);
}

function MakeSharingGroupPopup( title, ou, profileId, profileType, opener ){
	var popup = new D2L.Dialog();

	popup.SetTitle(new D2L.LP.Text.LangTerm('eP_PermissionProfiles.NewEdit.titSettings', title));
	popup.SetSize( '60%', '80%' );
	popup.SetSrc( '/d2l/ep/common/sharing/sharinggroup_newedit.d2l' );
	popup.SetSrcParam( 'ou', ou );
	popup.SetHasCloseIcon(true);

	popup.AddCustomButton(new D2L.LP.Text.LangTerm('Standard.Buttons.btnSaveAndClose'), D2L.Dialog.ButtonPosition.Left, D2L.Dialog.ResponseType.Positive, '1', 'CheckForAutoShareChanges', true);
	popup.SetSrcParam( 'profileId', profileId );
	popup.SetSrcParam( 'profileType', profileType );
	popup.SetCallback(ReloadPage);
	popup.Open( opener );
}

function ResetExternalInvite( shareableItem, contextId, onUpdateCallback ) {
	D2L.Rpc.Create(
			'ResetExternalInvite',
			onUpdateCallback,
			'/d2l/ep/ePObjects/sharing/rpc/rpc_functions.d2l'
		).Call( shareableItem, contextId );
}

function OpenIgnoreUserDialog( ignoredUserId, orgUnitId, openerId, contextId ) {
	// Open an MVC dialog
	var url = '/d2l/ep/' + orgUnitId + '/ignoreUser/ignoreUserDialogLaunch/' + ignoredUserId;

	var evt = D2L.LP.Web.UI.Desktop.MasterPages.Dialog.Open(
		new D2L.LP.Web.UI.Html.LegacyIdAdapter( openerId, contextId ),
		new D2L.LP.Web.Http.UrlLocation( url )
	);

	evt.AddListener( function( dontShowAgain ) {
		var n = new D2L.NavInfo();
		n.SetNavigation( '/d2l/ep/' + orgUnitId + '/explore/index' );
		Nav.Go( n );
	} );
}

function IgnoreUser( ignoredUserId, orgUnitId ) {
	// Do an MVC RPC call to ignore the user
	var pattern = "/d2l/ep/{orgUnitId}/ignoreUser/{action}/{id}";
	var routeVars = {
			'orgUnitId': orgUnitId,
			'action': 'ignoreUser',
			'id': ignoredUserId
		};

	var callback = function() {
		var n = new D2L.NavInfo();
		n.SetNavigation( '/d2l/ep/' + orgUnitId + '/explore/index' );
		Nav.Go( n );
	}

	D2L.LP.Web.UI.Rpc.Connect(
		D2L.LP.Web.UI.Rpc.Verbs.POST,
		new D2L.LP.Web.Mvc.RouteLocation.Create( pattern, null, routeVars ),
		{},
		{ success: callback }
	);
}

function ReloadPage( response ){
	Nav.Reload();
}


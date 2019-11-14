import * as React from "react";
import { useState, useEffect, useContext, useRef } from "react";
import { initializeIcons } from 'office-ui-fabric-react'
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardTitle,
    DocumentCardDetails,
    DocumentCardImage,
    IDocumentCardStyles,
    IDocumentCardActivityPerson,
    DocumentCardActions
  } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit, Image } from 'office-ui-fabric-react/lib/Image';
import { ID365DocumentCardProps } from './interface';
import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class Downloadfile implements ComponentFramework.FileObject{
    fileContent: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
}

const cardStyles: IDocumentCardStyles = {
  root: { cursor:"pointer",display: 'inline-block', marginRight: 20, marginBottom: 20, width: 200 }
};

export class D365CornerDocumentCard {
	
  attachmentId: string;
  name: string;
  size: number;
  extension: string;
  entityType: string;
  userImageUrl :string;
  userImageInitials: string;
  userFullName: string;
  lastModifiedOn : string;
  description :string;
  documentCardActivityPerson: IDocumentCardActivityPerson;
  documentImageSrcName: string;
  documentActivityPerson : IDocumentCardActivityPerson[];
  title:string;
  logoColor: string;
  constructor(title: string,attachmentId: string, description: string,  name: string, extension: string, size: number, userFullName: string, lastModifiedOn: string, userImageUrl?: string) {
      this.attachmentId = attachmentId;
      this.name = name;
      this.title = title ? title : "";
  this.size = size;
  this.description = description;
  this.extension = extension;
  this.userImageUrl = userImageUrl ? userImageUrl : "";
  this.userFullName = userFullName ? userFullName : "";
  this.lastModifiedOn = lastModifiedOn ? this.parseDateTime(lastModifiedOn) : "";
  this.userImageInitials = this.getInitials(this.userFullName);

  //Get Document Activity Person
  this.documentCardActivityPerson = { name: this.userFullName, profileImageSrc: this.userImageUrl, initials: this.userImageInitials };

  //Get Document Card Image
  this.documentImageSrcName = this.getDocumetCardImageName(this.extension);

  this.logoColor = this.getLogoColor(this.documentImageSrcName);

  this.documentActivityPerson = this.getDocActivityPerson(this.userFullName,this.userImageInitials, this.userImageUrl);
}

getDocActivityPerson(userFullName: string, userImageInitials: string, userImageUrl: string): IDocumentCardActivityPerson[] {
  
  const _temp: IDocumentCardActivityPerson[] = [
    { name : userFullName, profileImageSrc: userImageUrl, initials : userImageInitials }
  ];

  return _temp;
}

parseDateTime(lastModifiedOn: string): string {
  var dtStr:string ;

  var dt = new Date(lastModifiedOn);

  var month_names =["Jan","Feb","Mar",
                    "Apr","May","Jun",
                    "Jul","Aug","Sep",
                    "Oct","Nov","Dec"];
  
  var day = dt.getDate();
  var month_index = dt.getMonth();
  var year = dt.getFullYear();
  
  return "Modified "+ month_names[month_index] + " " + day + ", " + year;
}

public getLogoColor(source: string) : string {
  switch (source)
  {
     case "PDF" : return "#ce2c00";
     case "WordDocument" : return "#103f91";
     case "OneNoteLogo" : return "#813a7c";
     case "ExcelLogo" : return "#185c37";
     case "ImageDiff" : return "#813a7c";
     case "Video" : return "#813a7c";
     case "OutlookLogo" : return "#103f91";
     case "PowerPointDocument" : return "#ce2c00"
     default: return "#3b3a39"
  }
  
}


public getDocumetCardImageName(extension: string): string {
  switch (extension.toLowerCase())
  {
     case "pdf" : return "PDF";
     case "doc" : return "WordDocument";
     case "docx"  : return "WordDocument";
     case "one" : return "OneNoteLogo";
     case "xls" : return "ExcelLogo";
     case "xlsx" : return "ExcelLogo";
     case "jpeg" : return "ImageDiff";
     case "png" : return "ImageDiff";
     case "mkv" : return "Video";
     case "mp4" : return "Video";
     case "msg" : return "OutlookLogo";
     case "pptx" : return "PowerPointDocument";
     default: return "QuickNote"
  }
}


public getInitials(valStr: string) : string {
  var words = valStr.split(" "),
    initials = "";
  words.forEach(function(word) {
    initials += word.charAt(0);
  });
  return initials.toUpperCase();
}

public getExtensionFromMimeType(extension: string): string {
  return extension.split('/')[1];
} 

public onClickAttachment(id: string, _context: ComponentFramework.Context<IInputs>): void { 
      this.downloadAttachment(id, _context).then(f => {
          _context.navigation.openFile(f);   }
      );        
}

private downloadAttachment(id: string,_context: ComponentFramework.Context<IInputs>): 	Promise<Downloadfile> {
      debugger;
      return _context.webAPI.retrieveRecord("annotation", id).then(
          function success(result) {
              let file: Downloadfile = new Downloadfile();
              file.fileContent = result["documentbody"];
              file.fileName = result["filename"];
              file.fileSize = result["filesize"];
              file.mimeType = result["mimetype"];
              return file;               
          });
  }
}

const DocCardsApp: React.SFC<ID365DocumentCardProps> = (props) : JSX.Element => {
  return (
    <div>
      {
        props.documentCards.map(docCard => (
          <DocumentCard arial-label={docCard.name} styles={cardStyles} id=  {docCard.attachmentId} onClick={e => docCard.onClickAttachment(docCard.attachmentId, props._context) } >
            <DocumentCardImage height={150} imageFit={ImageFit.cover} iconProps={
              {  iconName: docCard.documentImageSrcName, styles: {
                  root: { color: docCard.logoColor , fontSize: '120px', width: '120px', height: '120px' } }
              } }
            />
            <DocumentCardDetails>
              <DocumentCardTitle title={docCard.title ? docCard.title : docCard.name} shouldTruncate />
            </DocumentCardDetails> 
            <DocumentCardActivity activity={docCard.lastModifiedOn} people={ docCard.documentActivityPerson } />
          </DocumentCard>
        ))
      }    
    </div>
  );
}

export { DocCardsApp };
{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice.DrawImage.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md }

{ ******************************************************* }
unit uOpenOffice.DrawImage;

interface

uses
  classes;

type
  TOpenOfficeDrawImage = class(TCollectionItem)
   private
    FUrlImage: string;
    FPositionX: integer;
    FPositionY: integer;
    FWidth: Integer;
    FHeight: Integer;

   public
     property URLImage: string read FUrlImage write FUrlImage;
     property PositionX: integer read FPositionX write FPositionX;
     property PositionY: integer read FPositionY write FPositionY;
     property Width: Integer read FWidth write FWidth;
     property Height: Integer read FHeight write FHeight;
  end;

implementation

end.

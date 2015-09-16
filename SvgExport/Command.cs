#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BoundarySegment = Autodesk.Revit.DB.BoundarySegment;
#endregion

namespace SvgExport
{
  [Transaction( TransactionMode.ReadOnly )]
  public class Command : IExternalCommand
  {
    const int _target_square_size = 100;

    /// <summary>
    /// Allow selection of room elements only.
    /// </summary>
    class RoomSelectionFilter : ISelectionFilter
    {
      public bool AllowElement( Element e )
      {
        return e is Room;
      }

      public bool AllowReference( Reference r, XYZ p )
      {
        return true;
      }
    }

    /// <summary>
    /// Get a room from the given document.
    /// If there is only one single room in the entire
    /// model, return that with no further ado. 
    /// Otherwise, if any elements have been pre-selected,
    /// pick the first room encountered among them.
    /// Otherwise, prompt the user to interactively 
    /// select a room or a room tag.
    /// </summary>
    Result GetRoom( UIDocument uidoc, out Room room )
    {
      room = null;

      Document doc = uidoc.Document;

      // Get all rooms in the model.

      FilteredElementCollector rooms
        = new FilteredElementCollector( doc )
          .OfClass( typeof( SpatialElement ) );

      if( 1 == rooms.Count()
        && rooms.FirstElement() is Room )
      {
        // If there is only one spatial element
        // and that is a room, pick that.

        room = rooms.FirstElement() as Room;
      }
      else
      {
        Selection sel = uidoc.Selection;

        // Check the preselacted elements.

        ICollection<ElementId> ids
          = sel.GetElementIds();

        foreach( ElementId id in ids )
        {
          Element e = doc.GetElement( id );
          if( e is Room )
          {
            room = e as Room;
            break;
          }
          if( e is RoomTag )
          {
            room = ( e as RoomTag ).Room;
            break;
          }
        }

        if( null == room )
        {
          // Prompt for interactive selection.

          try
          {
            Reference r = sel.PickObject(
              ObjectType.Element,
              new RoomSelectionFilter(),
              "Please select pick a room" );

            room = doc.GetElement( r.ElementId )
              as Room;
          }
          catch( Autodesk.Revit.Exceptions
            .OperationCanceledException )
          {
            return Result.Cancelled;
          }
        }
      }
      return Result.Succeeded;
    }

    /// <summary>
    /// Return an SVG representation of the
    /// given XYZ point scaled, offset and
    /// Y flipped to the target square size.
    /// </summary>
    string GetSvgPointFrom(
      XYZ p,
      XYZ pmid,
      double scale )
    {
      p -= pmid;
      p *= scale;
      int x = (int) ( p.X + 0.5 );
      int y = (int) ( p.Y + 0.5 );

      // The Revit Y coordinate points upwards,
      // the SVG one down, so flip the Y coord.

      y = -y;

      x += _target_square_size / 2;
      y += _target_square_size / 2;
      return x.ToString() + " " + y.ToString();
    }

    /// <summary>
    /// Generate and return an SVG path definition to
    /// represent the given room boundary loop, scaled 
    /// from the given bounding box to fit into a 
    /// 100 x 100 canvas. Actually, the size is 
    /// determined by _target_square_size.
    /// </summary>
    string GetSvgPathFrom(
      BoundingBoxXYZ bb,
      IList<BoundarySegment> loop )
    {
      // Determine scaling and offsets to transform
      // from bounding box to (0,0)-(100,100).

      XYZ pmin = bb.Min;
      XYZ pmax = bb.Max;
      XYZ vsize = pmax - pmin;
      XYZ pmid = pmin + 0.5 * vsize;
      double size = Math.Max( vsize.X, vsize.Y );
      double scale = _target_square_size / size;

      StringBuilder s = new StringBuilder();

      int nSegments = loop.Count;

      XYZ p0 = null; // loop start point
      XYZ p; // segment start point
      XYZ q = null; // segment end point

      foreach( BoundarySegment seg in loop )
      {
        Curve curve = seg.GetCurve();

        // Todo: handle non-linear curve.
        // Especially: if two long lines have a 
        // short arc in between them, skip the arc
        // and extend both lines.

        p = curve.GetEndPoint( 0 );

        Debug.Assert( null == q || q.IsAlmostEqualTo( p ),
          "expected last endpoint to equal current start point" );

        q = curve.GetEndPoint( 1 );

        if( null == p0 )
        {
          p0 = p; // save loop start point

          s.Append( "M"
            + GetSvgPointFrom( p, pmid, scale ) );
        }
        s.Append( "L"
          + GetSvgPointFrom( q, pmid, scale ) );
      }
      s.Append( "Z" );

      Debug.Assert( q.IsAlmostEqualTo( p0 ),
        "expected last endpoint to equal loop start point" );

      return s.ToString();
    }

    /// <summary>
    /// Invoke the SVG node.js web server.
    /// Use a local or global base URL and append
    /// the SVG path definition as a query string.
    /// Compare this with the JavaScript version used in
    /// http://the3dwebcoder.typepad.com/blog/2015/04/displaying-2d-graphics-via-a-node-server.html
    /// </summary>
    void DisplaySvg( string path_data )
    {
      var local = false;

      var base_url = local
        ? "http://127.0.0.1:5000"
        : "https://shielded-hamlet-1585.herokuapp.com";

      var d = path_data.Replace( ' ', '+' );

      var query_string = "d=" + d;

      string url = base_url + '?' + query_string;

      System.Diagnostics.Process.Start( url );
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      Room room = null;
      Result rc = GetRoom( uidoc, out room );

      SpatialElementBoundaryOptions opt
        = new SpatialElementBoundaryOptions();

      opt.SpatialElementBoundaryLocation =
        SpatialElementBoundaryLocation.Center; // loops closed

      //SpatialElementBoundaryLocation.Finish; // loops not closed

      IList<IList<BoundarySegment>> loops
        = room.GetBoundarySegments( opt );

      int nLoops = loops.Count;

      BoundingBoxXYZ bb = room.get_BoundingBox( null );

      string path_data = GetSvgPathFrom( bb, loops[0] );

      DisplaySvg( path_data );

      return Result.Succeeded;
    }
  }
}

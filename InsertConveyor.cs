using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Teigha.Runtime;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using BApp = Bricscad.ApplicationServices;

namespace Alfatech.FlexCAD
{

    internal class InsertConveyor {

        static bool needReferHeight(double h) { return h==0; }

        static (Conveyor,Matrix3d,Anchor?) jig(Conveyor conv,ObjectId convID)
        {
            var jigblk = new ComponentJigBlockReference(convID);
            var jig = new ConveyorJig(conv,jigblk);
            var drg_conv = jig.DragConveyor();
            var mtx = Matrix3d.Displacement(jigblk.BlockReference.Position - EntityUntil.GetLocation(convID));
            var anchor = jig.SelectedAnchor;
            return (drg_conv,mtx,anchor);
        }

        static double? referHeight()
        {
            var res = ReferConveyorHeight.SelectConveyor();
            if (res==null) {
                return null;
            }
            var (refH,gotH) = res.Value;
            if (gotH) {
                return refH;
            } else {
                return null;
            }
        }

        static void continuousInsert(Conveyor conv,ObjectId convID,Anchor anchor)
        {
            foreach (var id in ContinuousConveyor.Insert(convID,anchor)) {
                if (! id.IsNull) {
                    conv.SaveToComponent(id);
                }
            }
        }

        static void showMsg(string msg)
        {
            MessageBox.Show(msg,"FlexCAD",MessageBoxButtons.OK,MessageBoxIcon.Error);
        }

        static void applyMatrix(ObjectId id,Matrix3d mtx)
        {
            if (id.IsNull)
                return;
            using (var tr = id.Database.TransactionManager.StartTransaction()) {
                var ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                ent.TransformBy(mtx);
                tr.Commit();
            }
        }
        // returns: bool(keep going?)
        static bool _Cmd()
        {
            var conv = Conveyor.MakeConveyorByForm();
            if (conv==null) {
                return false;
            }
            var path = Alfatech.General.ComponentPath.Find(conv.ShortName);
            if (string.IsNullOrEmpty(path)) {
                showMsg($"選択された機種のコンベヤ図面 {conv.ShortName}.dwg が見つかりません。 ");
                return false;
            }
            var pos = Point3d.Origin;
            ObjectId convID;
            try {
                convID = InsertComponent.Insert(path,pos);
                EntityUntil.SetVisibility(convID,false);
            } catch (ArgumentNullException) {
                showMsg($"{conv.ShortName} のパスが空です");
                return true;
            } catch (FileNotFoundException) {
                showMsg($"{conv.ShortName} のパス {path} が存在しません");
                return true;
            }
            if (convID.IsNull) {
                return true;
            }
            var h = conv.SettingHeight;
            if (needReferHeight(h)) {
                var refh = referHeight();
                if (refh==null) {
                    EntityUntil.DeleteEntity(convID);
                    return true;
                }
                h = refh.Value;
            }
            var (drg_conv,mtx,anchor_p) = jig(conv,convID);
            if (drg_conv==null) {
                EntityUntil.DeleteEntity(convID);
                return true;
            }
            applyMatrix(convID,mtx);
            saveHeight(h.ToString());
            drg_conv.SaveToComponent(convID);
            EntityUntil.SetVisibility(convID,true);
            var anchor = anchor_p==null ? Anchor.End : anchor_p.Value;
            continuousInsert(drg_conv,convID,anchor);
            return true;
        }

        internal static void Cmd() {
            while (_Cmd()) {}
        }

    }

}

//#define DEBUG_PRINT
//#define ALLTERNATIVE_IMPL
using System;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System.Linq.Expressions;
using System.Drawing;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(UtilsZ.AdditionalCommands))]

namespace UtilsZ
{
	public class AdditionalCommands
	{

		[CommandMethod("AddPointZ")]
		public void AddPointZ()
		{
			Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
			PromptPointOptions getStartPoint = new PromptPointOptions("Select start point");
			PromptPointResult startPoint = ed.GetPoint(getStartPoint);

			PromptPointOptions getEndtPoint = new PromptPointOptions("Select end point");
			PromptPointResult endPoint = ed.GetPoint(getEndtPoint);

			PromptPointOptions getNewPoint = new PromptPointOptions("Select new point");
			PromptPointResult newPoint = ed.GetPoint(getNewPoint);

			//Create new point
			Database curDwg = Application.DocumentManager.MdiActiveDocument.Database;
			Transaction transaction = curDwg.TransactionManager.StartTransaction();
			BlockTable blkTbl;
			blkTbl = transaction.GetObject(curDwg.BlockTableId, OpenMode.ForRead) as BlockTable;

			BlockTableRecord blkTblRec;
			blkTblRec = transaction.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

			// Create a point with selected new point
			double xStartPoint = startPoint.Value.X;
			double yStartPoint = startPoint.Value.Y;
			double zStartPoint = startPoint.Value.Z;

			double xEndPoint = endPoint.Value.X;
			double yEndPoint = endPoint.Value.Y;
			double zEndPoint = endPoint.Value.Z;

			double xCreatePoint = newPoint.Value.X;
			double yCreatePoint = newPoint.Value.Y;
			double zCreatePoint;

			double distanceStartEndPoints = Math.Sqrt(Math.Pow((xStartPoint - xEndPoint), 2) + Math.Pow((yStartPoint - yEndPoint), 2));
			double distanceStartNewPoints = Math.Sqrt(Math.Pow((xStartPoint - xCreatePoint), 2) + Math.Pow((yStartPoint - yCreatePoint), 2));

			zCreatePoint = Math.Abs(zStartPoint - zEndPoint) * distanceStartNewPoints / distanceStartEndPoints + zStartPoint;

			//point addition TODO commented START
			DBPoint dbPoint = new DBPoint(new Point3d(xCreatePoint, yCreatePoint, zCreatePoint));
			dbPoint.SetDatabaseDefaults();
			// Add the new object to the block table record and the transaction
			blkTblRec.AppendEntity(dbPoint);
			transaction.AddNewlyCreatedDBObject(dbPoint, true);
			//point addition TODO commented END

			// Save the new object to the database
			transaction.Commit();

#if (DEBUG_PRINT)
			Application.ShowAlertDialog("start point = (" + xStartPoint.ToString() + " ," + yStartPoint.ToString() + " ," + zStartPoint.ToString() + ")");
			Application.ShowAlertDialog("end point = (" + xEndPoint.ToString() + " ," + yEndPoint.ToString() + " ," + zEndPoint.ToString() + ")");
			Application.ShowAlertDialog("distanceStartEndPoints=" + distanceStartEndPoints.ToString() + "; distanceStartNewPoints=" + distanceStartNewPoints.ToString() + ")");
#endif
		}
#if (NOT_WORKING)
		public static void AddPointZ()
		{
			string blockName = "utilsPoints";

			Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
			Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

			PromptPointOptions getStartPoint = new PromptPointOptions("Select start point");
			PromptPointResult startPoint = acDocEd.GetPoint(getStartPoint);

			PromptPointOptions getEndtPoint = new PromptPointOptions("Select end point");
			PromptPointResult endPoint = acDocEd.GetPoint(getEndtPoint);

			PromptPointOptions getNewPoint = new PromptPointOptions("Select new point");
			PromptPointResult newPoint = acDocEd.GetPoint(getNewPoint);

			using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
			{
				// Open the block table and check if it contains "BlockMtext"
				BlockTable bt = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
				ObjectId btrId = new ObjectId();
				BlockTableRecord btr = new BlockTableRecord();

				if (bt.Has(blockName))
				{
					btrId = bt[blockName];
					btr = (BlockTableRecord)acTrans.GetObject(btrId, OpenMode.ForRead);
				}
				// else, create a new block definition
				else
				{
					// Create a new block definition
					btr = new BlockTableRecord();
					btr.Name = blockName;

					// Add the block definition to the block table
					acTrans.GetObject(bt.ObjectId, OpenMode.ForWrite);
					btrId = bt.Add(btr);
					acTrans.AddNewlyCreatedDBObject(btr, true);
				}

				// Create a point with selected new point
				double xStartPoint = startPoint.Value.X;
				double yStartPoint = startPoint.Value.Y;
				double zStartPoint = startPoint.Value.Z;

				double xEndPoint = endPoint.Value.X;
				double yEndPoint = endPoint.Value.Y;
				double zEndPoint = endPoint.Value.Z;

				double xCreatePoint = newPoint.Value.X;
				double yCreatePoint = newPoint.Value.Y;
				double zCreatePoint;

				double distanceStartEndPoints = Math.Sqrt(Math.Pow((xStartPoint - xEndPoint), 2) + Math.Pow((yStartPoint - yEndPoint), 2));
				double distanceStartNewPoints = Math.Sqrt(Math.Pow((xStartPoint - xCreatePoint), 2) + Math.Pow((yStartPoint - yCreatePoint), 2));

				zCreatePoint = Math.Abs(zStartPoint - zEndPoint) * distanceStartNewPoints / distanceStartEndPoints + zStartPoint;

				Point3d createdPoint = new Point3d(xCreatePoint, yCreatePoint, zCreatePoint);
				BlockReference br = new BlockReference(createdPoint, btrId);

				BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

				currentSpace.AppendEntity(br);
				acTrans.AddNewlyCreatedDBObject(br, true);
				acTrans.Commit();
			}
		}
#endif
		[CommandMethod("ChangeNameBR")]
		public static void ChangeNameBR()
		{
			// Get the current document editor
			Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
			Database db = Application.DocumentManager.MdiActiveDocument.Database;
			// Create a TypedValue array to define the filter criteria
			TypedValue[] acTypValAr = new TypedValue[1];
			acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 0);
			// Assign the filter criteria to a SelectionFilter object
			SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
			// Request for objects to be selected in the drawing area
			PromptSelectionResult acSSPrompt = acDocEd.GetSelection(acSelFtr);
			// If the prompt status is OK, objects were selected
			if (acSSPrompt.Status == PromptStatus.OK)
			{
				SelectionSet acSSet = acSSPrompt.Value;
				SelectedObject[] chosenMtextArr = new SelectedObject[1];
				acSSet.CopyTo(chosenMtextArr, 0);
				Transaction trans = db.TransactionManager.StartTransaction();
				MText mtext = trans.GetObject(chosenMtextArr[0].ObjectId, OpenMode.ForRead, true) as MText;
				string mtextText = mtext.Contents.ToString();
				Application.ShowAlertDialog("Selected object: " + mtextText);

				//Block reference
				PromptEntityResult promptEntity = acDocEd.GetEntity("Choose a Block Reference");
				if (acSSPrompt.Status != PromptStatus.OK)
				{
					acDocEd.WriteMessage("Block Reference choosing failed!");
					return;
				}

				string attbName = "HEIGHT";
				BlockReference blockReference = trans.GetObject(promptEntity.ObjectId, OpenMode.ForRead, true) as BlockReference;
				foreach (ObjectId arId in blockReference.AttributeCollection)
				{
					AttributeReference ar = trans.GetObject(arId, OpenMode.ForRead) as AttributeReference;
					if (null == ar)
					{
						acDocEd.WriteMessage("AttributeReference getting failed!");
						return;
					}
					if (ar.Tag.ToUpper() == attbName)
					{
						ar.UpgradeOpen();
						ar.TextString = mtextText;
						ar.DowngradeOpen();
					}
				}

				trans.Commit();
			}
			else
			{
				Application.ShowAlertDialog("Number of objects selected: 0");
			}
		}

		/// <summary>
		/// Creates from each mtext in layers: M1506_T and M1507_T a blockreference in layers: M1506 and M1507 respectively.  
		/// </summary>
		[CommandMethod("BRfromMtext")]
		public static void BRfromMtext()
		{
			string attbName = "HEIGHT";
			string[] btrWantedNames = new string[2] { "M1506", "M1507" };
			string mtextLayerSuffix = "_T";
			//string blockName = "BlockMtext";

			// Get the current docusment and database, and start a transaction
			Document acDoc = Application.DocumentManager.MdiActiveDocument;
			Database acCurDb = acDoc.Database;
			Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;

			using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
			{
				// Open the block table and check if it contains "BlockMtext"
				BlockTable bt = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
				ObjectId[] btrId = new ObjectId[btrWantedNames.Length];
				BlockTableRecord[] btr = new BlockTableRecord[btrWantedNames.Length];
				int i = 0;
				// if the block table contains "BlockMtext", get its ObjectId and open it
				foreach (string blockName in btrWantedNames)
				{
					if (bt.Has(blockName))
					{
						btrId[i] = bt[blockName];
						btr[i] = (BlockTableRecord)acTrans.GetObject(btrId[i], OpenMode.ForRead);
					}
					// else, create a new block definition
					else
					{
						// Create a new block definition
						btr[i] = new BlockTableRecord();
						btr[i].Name = blockName;

						// Add the block definition to the block table
						acTrans.GetObject(bt.ObjectId, OpenMode.ForWrite);
						btrId[i] = bt.Add(btr[i]);
						acTrans.AddNewlyCreatedDBObject(btr[i], true);

						// Add an attribute definition to the block definition
						var attDef = new AttributeDefinition(Point3d.Origin, "---", attbName, "", acCurDb.Textstyle);
						//attDef.IsMTextAttributeDefinition = true;
						btr[i].AppendEntity(attDef);
						acTrans.AddNewlyCreatedDBObject(attDef, true);
					}
					i++;
				}
				// For each MText found in the current space, insert the block reference
				BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
				foreach (ObjectId objId in currentSpace)
				{
					if ((objId.ObjectClass.DxfName == "MTEXT")||(objId.ObjectClass.DxfName == "TEXT"))
					{

						string brLayerName;
						Point3d location;
						double rotation;
						string contents;

						if (objId.ObjectClass.DxfName == "MTEXT")
						{
							// Open the MText
							MText mtext = acTrans.GetObject(objId, OpenMode.ForRead, true) as MText;
							brLayerName = mtext.Layer.Substring(0, (mtext.Layer.Length - mtextLayerSuffix.Length));
							location = mtext.Location;
							rotation = mtext.Rotation;
							contents = mtext.Contents;
						}
						else
						{
							DBText dbText = acTrans.GetObject(objId, OpenMode.ForRead, true) as DBText;
							brLayerName = "M" + dbText.Layer; // TEXT is in layer called 1506 & 1507, without M
							location = dbText.Position;
							rotation = dbText.Rotation;
							contents = dbText.TextString;
						}
															
						i = 0;
						foreach (string blockName in btrWantedNames)
						{
							if (brLayerName == blockName)
							{
								break;
							}
							i++;
						}
						
						if(i >= btrWantedNames.Length)
						{
							//TEXT or MTEXT is in differnet layer then:
							//MTEXT: M1506_T, M1507_T
							//TEXT: 1560, 1507
							continue;
						}
						//Insert a block reference
						BlockReference br = new BlockReference(location, btrId[i]);
						currentSpace.AppendEntity(br);
						acTrans.AddNewlyCreatedDBObject(br, true);
						br.Layer = brLayerName;
						br.Rotation = rotation;
						// Add the attribute references to the attribute collection of the block reference
						foreach (ObjectId attId in btr[i])
						{
							if (attId.ObjectClass.DxfName == "ATTDEF")
							{
								var attDef = (AttributeDefinition)acTrans.GetObject(attId, OpenMode.ForRead);
								var attRef = new AttributeReference();
								attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
								if (attDef.Tag.ToUpper() == attbName)
								{
									attRef.TextString = contents;
								}
								br.AttributeCollection.AppendAttribute(attRef);
								acTrans.AddNewlyCreatedDBObject(attRef, true);
							}
						}
					}
				}
				acTrans.Commit();
			}
		}
	}
}

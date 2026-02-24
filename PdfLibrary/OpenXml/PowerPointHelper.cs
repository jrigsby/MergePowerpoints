using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

namespace PdfLibrary.OpenXml;

public class PowerPointHelper {
    private uint _uniqueId;

    private static uint GetMaxSlideMasterId(SlideMasterIdList? slideMasterIdList) {
        uint max = 2147483648;
        if (slideMasterIdList != null) {
            foreach (SlideMasterId child in slideMasterIdList.Elements<SlideMasterId>()) {
                uint id = child.Id;
                if (id > max)
                    max = id;
            }
        }

        return max;
    }

    private static uint GetMaxSlideId(SlideIdList? slideIdList) {
        uint max = 256;
        if (slideIdList != null) {
            foreach (SlideId child in slideIdList.Elements<SlideId>()) {
                uint id = child.Id ?? throw new InvalidOperationException();
                if (id > max)
                    max = id;
            }
        }

        return max;
    }

    private void FixSlideLayoutIds(PresentationPart presPart) {
        foreach (var slideMasterPart in presPart.SlideMasterParts) {
            foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList) {
                _uniqueId++;
                slideLayoutId.Id = _uniqueId;
            }

            slideMasterPart.SlideMaster.Save();
        }
    }

    public static string FormatErrors(IEnumerable<ValidationErrorInfo> errors) {
        var errorIndex = 1;
        var builder = new StringBuilder();

        var validationErrorInfos = errors.ToList();
        foreach (ValidationErrorInfo errorInfo in validationErrorInfos) {
            builder.Append(errorInfo.Description);
            builder.Append('\n');
            builder.Append(errorInfo.Path?.XPath);
            if (++errorIndex <= validationErrorInfos.Count())
                builder.Append("-------------------");
        }

        return builder.ToString();
    }

    public void MergePresentationSlides(string sourceFolderLocation, string sourcePresentation,
        string destinationFolderLocation, string destPresentation) {
        var id = 0;
        using (var myDestDeck =
               PresentationDocument.Open(destinationFolderLocation + destPresentation, true)) {
            var destPresPart = myDestDeck.PresentationPart;
            if (destPresPart.Presentation.SlideIdList == null)
                destPresPart.Presentation.SlideIdList = new SlideIdList();

            using (var mySourceDeck =
                   PresentationDocument.Open(sourceFolderLocation + sourcePresentation, false)) {
                var sourcePresPart = mySourceDeck.PresentationPart;

                _uniqueId = GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);
                var maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);
                foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList) {
                    id++;
                    var sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
                    var relId = sourcePresentation.Remove(sourcePresentation.IndexOf('.')) + id;
                    var destSp = destPresPart.AddPart(sp, relId);
                    var destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                    destPresPart.AddPart(destMasterPart);

                    _uniqueId++;
                    var newSlideMasterId = new SlideMasterId {
                        RelationshipId = destPresPart.GetIdOfPart(destMasterPart),
                        Id = _uniqueId
                    };

                    maxSlideId++;
                    var newSlideId = new SlideId {
                        RelationshipId = relId,
                        Id = maxSlideId
                    };
                    destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                    destPresPart.Presentation.SlideIdList.Append(newSlideId);
                }

                FixSlideLayoutIds(destPresPart);
            }

            destPresPart.Presentation.Save();
        }
    }

    public void MergePresentationSlidesStream(PresentationDocument mySourceDeck, string sourceName,
        string destinationFolderLocation, string destPresentation) {
        int id = 0;
        using (PresentationDocument myDestDeck =
               PresentationDocument.Open(destinationFolderLocation + destPresentation, true)) {
            var destPresPart = myDestDeck.PresentationPart;
            if (destPresPart.Presentation.SlideIdList == null)
                destPresPart.Presentation.SlideIdList = new SlideIdList();

            // using (PresentationDocument mySourceDeck =
            //        PresentationDocument.Open(sourceFolderLocation + sourcePresentation, false)) {
            PresentationPart sourcePresPart = mySourceDeck.PresentationPart;

            _uniqueId = GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);
            uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);
            foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList) {
                id++;
                var sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
                var relId = sourceName.Remove(sourceName.IndexOf('.')) + id;
                var destSp = destPresPart.AddPart(sp, relId);
                var destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                destPresPart.AddPart(destMasterPart);

                _uniqueId++;
                var newSlideMasterId = new SlideMasterId();
                newSlideMasterId.RelationshipId = destPresPart.GetIdOfPart(destMasterPart);
                newSlideMasterId.Id = _uniqueId;

                maxSlideId++;
                var newSlideId = new SlideId {
                    RelationshipId = relId,
                    Id = maxSlideId
                };
                destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                destPresPart.Presentation.SlideIdList.Append(newSlideId);
            }

            FixSlideLayoutIds(destPresPart);

            destPresPart.Presentation.Save();
        }
    }

    public void MergePresentationSlidesStreams(PresentationDocument mySourceDeck, string sourceName,
        PresentationDocument myDestDeck) {
        int id = 0;
        if (myDestDeck.PresentationPart == null) throw new Exception("expected PresentationPart, found null");
        if (myDestDeck.PresentationPart == null) throw new Exception("expected destPresPart.Presentation, found null");
        
        PresentationPart destPresPart = myDestDeck.PresentationPart;
        if (destPresPart.Presentation == null) throw new Exception("expected destPresPart.Presentation, found null");
        
        if (destPresPart.Presentation.SlideIdList == null)
            destPresPart.Presentation.SlideIdList = new SlideIdList();

        // using (PresentationDocument mySourceDeck =
        //        PresentationDocument.Open(sourceFolderLocation + sourcePresentation, false)) {
        if (mySourceDeck.PresentationPart == null) throw new Exception("expected mySourceDeck.PresentationPart, found null");
        PresentationPart sourcePresPart = mySourceDeck.PresentationPart;
        if (sourcePresPart.Presentation == null) throw new Exception("expected sourcePresPart.Presentation, found null");
        if (sourcePresPart.Presentation.SlideIdList == null) throw new Exception("expected sourcePresPart.Presentation.SlideIdList, found null");
        
        _uniqueId = GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);
        uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);
        foreach (var openXmlElement in sourcePresPart.Presentation.SlideIdList) {
            var slideId = (SlideId)openXmlElement;
            id++;
            var sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
            //var relId = sourceName.Remove(sourceName.IndexOf('.')) + id; //orioginal working code I used $"big-{i}.pptx" as sourceName
            var relId = "_" + sourceName + "_" + id;  //sourceName is a guid in my test app, so this works. Follows rules for XSD ID, cannot start with number, guid can, but can start with underscore.
            var destSp = destPresPart.AddPart(sp, relId);
            var destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
            destPresPart.AddPart(destMasterPart);

            _uniqueId++;
            var newSlideMasterId = new SlideMasterId {
                RelationshipId = destPresPart.GetIdOfPart(destMasterPart),
                Id = _uniqueId
            };

            maxSlideId++;
            var newSlideId = new SlideId {
                RelationshipId = relId,
                Id = maxSlideId
            };
            destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
            destPresPart.Presentation.SlideIdList.Append(newSlideId);
        }

        FixSlideLayoutIds(destPresPart);

        destPresPart.Presentation.Save();
    }
}
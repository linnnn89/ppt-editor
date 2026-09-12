These five XML files are the organization-chart data/layout/style/colors/frame
from the .NET Foundation Open XML SDK MIT-licensed test fixture:
https://github.com/dotnet/Open-XML-SDK/blob/main/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/presentation/SmartArt_OrgChart1.pptx

Upstream Git blob: 0b5c07abb660ffe67ce8250bcbce0bd447ade5e2 (69,565 bytes).
Only diagram XML was extracted. Document metadata, thumbnail, notes, animation
and source theme/master were not included. `../smartart-seed.js` wraps these
static XML parts in a newly generated synthetic presentation for regression
testing. Original labels are generic structural labels (Top, Bottom-Left, etc.).
Copyright and license are retained in LICENSE.txt. This is a test seed, not a
bundled network slide template or a claim that all SmartArt layouts are supported.

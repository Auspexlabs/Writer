// writer-vision: the Writer engine's picture helper on macOS. Background removal with Apple Vision (the framework behind
// Finder's "Remove Background"), re-encoding with ImageIO. The command line is in README.md next to this file.
import CoreImage
import Foundation
import ImageIO
import UniformTypeIdentifiers
import Vision

func fail(_ message: String, _ code: Int32 = 1) -> Never {
    FileHandle.standardError.write(Data((message + "\n").utf8))
    exit(code)
}

let usage = """
    usage: writer-vision cutout <in> <out.png> [--at x,y]
           writer-vision compress <in> <out.png|out.jpg> --max WxH [--crop l,t,r,b] [--quality q]
    """

/// The picture upright (its EXIF orientation applied), at full size.
func load(_ path: String) -> CGImage {
    guard let source = CGImageSourceCreateWithURL(URL(fileURLWithPath: path) as CFURL, nil),
          let props = CGImageSourceCopyPropertiesAtIndex(source, 0, nil) as? [CFString: Any],
          let width = props[kCGImagePropertyPixelWidth] as? Int, let height = props[kCGImagePropertyPixelHeight] as? Int
    else { fail("cannot read \(path)") }
    let options: [CFString: Any] = [
        kCGImageSourceCreateThumbnailFromImageAlways: true,
        kCGImageSourceCreateThumbnailWithTransform: true,
        kCGImageSourceThumbnailMaxPixelSize: max(width, height),
        kCGImageSourceShouldCacheImmediately: true,
    ]
    guard let image = CGImageSourceCreateThumbnailAtIndex(source, 0, options as CFDictionary) else { fail("cannot decode \(path)") }
    return image
}

func write(_ image: CGImage, to path: String) {
    let jpeg = [".jpg", ".jpeg"].contains { path.lowercased().hasSuffix($0) }
    let type = jpeg ? UTType.jpeg : UTType.png
    guard let destination = CGImageDestinationCreateWithURL(URL(fileURLWithPath: path) as CFURL, type.identifier as CFString, 1, nil)
    else { fail("cannot write \(path)") }
    let quality = Double(option("--quality") ?? "") ?? 0.85
    CGImageDestinationAddImage(destination, image, (jpeg ? [kCGImageDestinationLossyCompressionQuality: quality] : [:]) as CFDictionary)
    if !CGImageDestinationFinalize(destination) { fail("cannot write \(path)") }
}

func option(_ name: String) -> String? {
    let args = CommandLine.arguments
    guard let i = args.firstIndex(of: name), i + 1 < args.count else { return nil }
    return args[i + 1]
}

func numbers(_ text: String, _ count: Int, _ separator: Character) -> [Double] {
    let values = text.split(separator: separator).compactMap { Double($0.trimmingCharacters(in: .whitespaces)) }
    if values.count != count || values.contains(where: { !$0.isFinite }) { fail("'\(text)' is not valid\n\(usage)", 64) }
    return values
}

/// Every foreground instance Vision finds, or only the one under the point (0..1, from the top left).
func cutout(_ input: String, _ output: String) {
    let image = load(input)
    let request = VNGenerateForegroundInstanceMaskRequest()
    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do { try handler.perform([request]) } catch { fail("Vision failed: \(error.localizedDescription)") }
    guard let result = request.results?.first, !result.allInstances.isEmpty else { fail("no subject found", 2) }
    var instances = result.allInstances
    if let at = option("--at") {
        let p = numbers(at, 2, ",")
        let mask = result.instanceMask
        CVPixelBufferLockBaseAddress(mask, .readOnly)
        defer { CVPixelBufferUnlockBaseAddress(mask, .readOnly) }
        let w = CVPixelBufferGetWidth(mask), h = CVPixelBufferGetHeight(mask)
        let x = min(w - 1, max(0, Int(p[0] * Double(w)))), y = min(h - 1, max(0, Int(p[1] * Double(h))))
        guard let base = CVPixelBufferGetBaseAddress(mask) else { fail("Vision returned no mask") }
        let label = Int(base.load(fromByteOffset: y * CVPixelBufferGetBytesPerRow(mask) + x, as: UInt8.self))
        if label == 0 { fail("no subject at \(at)", 2) }
        instances = IndexSet(integer: label)
    }
    let masked: CVPixelBuffer
    do { masked = try result.generateMaskedImage(ofInstances: instances, from: handler, croppedToInstancesExtent: false) }
    catch { fail("Vision failed: \(error.localizedDescription)") }
    let picture = CIImage(cvPixelBuffer: masked)
    guard let cut = CIContext().createCGImage(picture, from: picture.extent) else { fail("cannot render the cut-out") }
    write(cut, to: output)
}

/// Cuts the given fractions off the edges, then scales down (never up) to fit within the box.
func compress(_ input: String, _ output: String) {
    var image = load(input)
    guard let max = option("--max") else { fail(usage, 64) }
    let box = numbers(max, 2, "x")
    if let crop = option("--crop") {
        let c = numbers(crop, 4, ",")
        let w = Double(image.width), h = Double(image.height)
        let rect = CGRect(x: (c[0] * w).rounded(), y: (c[1] * h).rounded(), width: ((1 - c[0] - c[2]) * w).rounded(), height: ((1 - c[1] - c[3]) * h).rounded())
        guard rect.width >= 1, rect.height >= 1, let cropped = image.cropping(to: rect) else { fail("crop \(crop) leaves nothing") }
        image = cropped
    }
    let scale = min(1, box[0] / Double(image.width), box[1] / Double(image.height))
    if scale < 1 {
        let w = Swift.max(1, Int((Double(image.width) * scale).rounded())), h = Swift.max(1, Int((Double(image.height) * scale).rounded()))
        let opaque = [.none, .noneSkipFirst, .noneSkipLast].contains(image.alphaInfo)
        guard let space = CGColorSpace(name: CGColorSpace.sRGB),
              let context = CGContext(data: nil, width: w, height: h, bitsPerComponent: 8, bytesPerRow: 0, space: space,
                                      bitmapInfo: (opaque ? CGImageAlphaInfo.noneSkipLast : CGImageAlphaInfo.premultipliedLast).rawValue)
        else { fail("cannot scale the picture") }
        context.interpolationQuality = .high
        context.draw(image, in: CGRect(x: 0, y: 0, width: w, height: h))
        guard let scaled = context.makeImage() else { fail("cannot scale the picture") }
        image = scaled
    }
    write(image, to: output)
}

let args = CommandLine.arguments
guard args.count >= 4 else { fail(usage, 64) }
switch args[1] {
case "cutout": cutout(args[2], args[3])
case "compress": compress(args[2], args[3])
default: fail(usage, 64)
}

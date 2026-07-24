import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

export const sha256 = value => crypto.createHash('sha256').update(value).digest('hex');

export const readJson = file => JSON.parse(fs.readFileSync(file, 'utf8'));

export const readJsonLines = file => fs.readFileSync(file, 'utf8')
  .split(/\r?\n/)
  .filter(Boolean)
  .map(line => JSON.parse(line));

export const readGzipJsonLines = file => zlib.gunzipSync(fs.readFileSync(file))
  .toString('utf8')
  .split(/\r?\n/)
  .filter(Boolean)
  .map(line => JSON.parse(line));

export const deterministicGzip = value => zlib.gzipSync(value, {
  level: zlib.constants.Z_BEST_COMPRESSION,
  mtime: 0
});

export const writeJson = (file, value, compact = false) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, `${JSON.stringify(value, null, compact ? 0 : 2)}\n`, 'utf8');
};

export const writeGzip = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, deterministicGzip(value));
};

export const copySourceFile = (sourceRoot, relativePath, outputRoot, gzip = true) => {
  const sourcePath = path.join(sourceRoot, ...relativePath.split('/'));
  const source = fs.readFileSync(sourcePath);
  const outputPath = path.join(
    outputRoot,
    ...`${relativePath}${gzip ? '.gz' : ''}`.split('/')
  );
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  const output = gzip ? deterministicGzip(source) : source;
  fs.writeFileSync(outputPath, output);
  return {
    path: relativePath,
    records: relativePath.endsWith('.jsonl')
      ? source.toString('utf8').split(/\r?\n/).filter(Boolean).length
      : null,
    sourceSha256: sha256(source),
    localPath: path.relative(outputRoot, outputPath).replaceAll('\\', '/'),
    localSha256: sha256(output),
    sourceBytes: source.length,
    localBytes: output.length,
    compression: gzip ? 'gzip' : 'none'
  };
};

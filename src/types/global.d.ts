// Node 18+ (VS Code's bundled runtime) ships global fetch and URL implementations.
// @types/node in this project predates both, so declare just enough to compile
// without bumping the whole (very old) TypeScript/@types/node toolchain.
declare function fetch(input: any, init?: any): Promise<any>;
declare class URL {
    constructor(input: string, base?: string);
    origin: string;
    href: string;
    pathname: string;
}

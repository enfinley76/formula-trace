// formula-trace answers one question: given a cell in a spreadsheet, what is
// the full chain of cells that feed into it (or that depend on it), and in
// what order would they need to be recalculated.
package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// cellRefPattern matches a single cell reference or a range of two, with
// optional $ anchors (A1, $A$1, B2:B9) and an optional leading sheet
// qualifier (Sheet2!A1, 'My Sheet'!A1:B2). The sheet qualifier, when
// present, only needs to appear before the first cell of a range - Excel
// doesn't allow "Sheet1!A1:Sheet2!B2".
var cellRefPattern = regexp.MustCompile(`(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?`)

// Grid holds every cell's raw formula text (empty for literals) and its
// resolved, deduplicated list of direct dependencies.
type Grid struct {
	formulas map[string]string
	deps     map[string][]string
}

// splitSheetPrefix separates a "Sheet1!A1" or "'My Sheet'!A1" style
// reference into its sheet name (quotes stripped) and the bare cell part.
// A reference with no "!" has no sheet and returns an empty sheet name.
func splitSheetPrefix(s string) (sheet, rest string) {
	i := strings.LastIndex(s, "!")
	if i == -1 {
		return "", s
	}
	return strings.Trim(s[:i], "'"), s[i+1:]
}

func normalizeCell(s string) string {
	sheet, cell := splitSheetPrefix(strings.TrimSpace(s))
	cell = strings.ToUpper(strings.ReplaceAll(cell, "$", ""))
	if sheet == "" {
		return cell
	}
	return sheet + "!" + cell
}

// splitCell separates a normalized reference like "AB12" into its column
// letters and row number.
func splitCell(ref string) (col string, row int, err error) {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		i++
	}
	if i == 0 || i == len(ref) {
		return "", 0, fmt.Errorf("not a cell reference: %q", ref)
	}
	row, err = strconv.Atoi(ref[i:])
	if err != nil {
		return "", 0, fmt.Errorf("not a cell reference: %q", ref)
	}
	return ref[:i], row, nil
}

func colToNum(col string) int {
	n := 0
	for _, c := range col {
		n = n*26 + int(c-'A'+1)
	}
	return n
}

func numToCol(n int) string {
	var b []byte
	for n > 0 {
		n--
		b = append([]byte{byte('A' + n%26)}, b...)
		n /= 26
	}
	return string(b)
}

// expandRange turns "A1:B3" style bounds into every individual cell in the
// rectangle, in row-major order.
func expandRange(from, to string) ([]string, error) {
	c1, r1, err := splitCell(from)
	if err != nil {
		return nil, err
	}
	c2, r2, err := splitCell(to)
	if err != nil {
		return nil, err
	}
	n1, n2 := colToNum(c1), colToNum(c2)
	if n1 > n2 {
		n1, n2 = n2, n1
	}
	if r1 > r2 {
		r1, r2 = r2, r1
	}
	var cells []string
	for r := r1; r <= r2; r++ {
		for c := n1; c <= n2; c++ {
			cells = append(cells, fmt.Sprintf("%s%d", numToCol(c), r))
		}
	}
	return cells, nil
}

// extractRefs pulls every cell reference out of a formula, expanding ranges
// and deduplicating while preserving first-seen order. A sheet qualifier on
// a range applies to every cell the range expands to.
func extractRefs(formula string) []string {
	matches := cellRefPattern.FindAllString(formula, -1)
	seen := make(map[string]bool)
	var out []string
	add := func(cell string) {
		if !seen[cell] {
			seen[cell] = true
			out = append(out, cell)
		}
	}
	for _, m := range matches {
		sheet, rest := splitSheetPrefix(m)
		rest = strings.ToUpper(strings.ReplaceAll(rest, "$", ""))
		if strings.Contains(rest, ":") {
			parts := strings.SplitN(rest, ":", 2)
			cells, err := expandRange(parts[0], parts[1])
			if err != nil {
				continue
			}
			for _, c := range cells {
				if sheet != "" {
					c = sheet + "!" + c
				}
				add(c)
			}
		} else {
			if sheet != "" {
				rest = sheet + "!" + rest
			}
			add(rest)
		}
	}
	return out
}

// loadGrid reads a CSV with a "cell,formula" header. A formula starting with
// "=" is parsed for dependencies; anything else is treated as a literal
// value with no dependencies.
func loadGrid(path string) (*Grid, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.FieldsPerRecord = -1

	header, err := r.Read()
	if err != nil {
		return nil, fmt.Errorf("reading header: %w", err)
	}
	if len(header) < 2 || strings.ToLower(strings.TrimSpace(header[0])) != "cell" {
		return nil, fmt.Errorf("expected header \"cell,formula\", got %v", header)
	}

	g := &Grid{formulas: make(map[string]string), deps: make(map[string][]string)}
	for {
		rec, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		if len(rec) < 2 {
			continue
		}
		cell := normalizeCell(strings.TrimSpace(rec[0]))
		if cell == "" {
			continue
		}
		formula := strings.TrimSpace(rec[1])
		g.formulas[cell] = formula
		if strings.HasPrefix(formula, "=") {
			g.deps[cell] = extractRefs(formula)
		}
	}
	return g, nil
}

const (
	unvisited = iota
	inProgress
	done
)

// walk does a depth-first traversal over adj starting at target, appending
// each node to order only after all of its edges have been visited, which
// yields a valid calculation order. A cycle is reported as an error with the
// path that closes the loop.
func walk(target string, adj map[string][]string) ([]string, error) {
	state := make(map[string]int)
	var order []string
	var path []string

	var visit func(cell string) error
	visit = func(cell string) error {
		switch state[cell] {
		case done:
			return nil
		case inProgress:
			cycle := append(append([]string{}, path...), cell)
			return fmt.Errorf("circular reference: %s", strings.Join(cycle, " -> "))
		}
		state[cell] = inProgress
		path = append(path, cell)
		for _, dep := range adj[cell] {
			if err := visit(dep); err != nil {
				return err
			}
		}
		path = path[:len(path)-1]
		state[cell] = done
		order = append(order, cell)
		return nil
	}

	if err := visit(target); err != nil {
		return nil, err
	}
	return order, nil
}

// reverse builds a dependents graph: for every cell -> deps edge in g, add
// dep -> cell in the result.
func reverseGraph(g *Grid) map[string][]string {
	rev := make(map[string][]string)
	for cell, deps := range g.deps {
		for _, dep := range deps {
			rev[dep] = append(rev[dep], cell)
		}
	}
	for k := range rev {
		sort.Strings(rev[k])
	}
	return rev
}

func main() {
	dependents := flag.Bool("dependents", false, "trace cells that depend on the target instead of cells it depends on")
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, "usage: %s [-dependents] <sheet.csv> <cell>\n", os.Args[0])
		flag.PrintDefaults()
	}
	flag.Parse()

	if flag.NArg() != 2 {
		flag.Usage()
		os.Exit(2)
	}
	path := flag.Arg(0)
	target := normalizeCell(flag.Arg(1))

	g, err := loadGrid(path)
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
	if _, ok := g.formulas[target]; !ok {
		fmt.Fprintf(os.Stderr, "error: cell %s not found in %s\n", target, path)
		os.Exit(1)
	}

	adj := g.deps
	if *dependents {
		adj = reverseGraph(g)
	}

	order, err := walk(target, adj)
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}

	for _, cell := range order {
		if cell == target {
			continue
		}
		formula, ok := g.formulas[cell]
		switch {
		case !ok:
			formula = "(not in sheet)"
		case formula == "":
			formula = "(literal)"
		}
		fmt.Printf("%s\t%s\n", cell, formula)
	}
}

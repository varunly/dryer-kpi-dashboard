"""
Simple Production Optimizer using Pre-built Database
Uses the optimization database to instantly find best production sequence
"""

import json
from itertools import permutations
import pandas as pd

class SimpleProductionOptimizer:
    def __init__(self, database_file="optimization_database.json"):
        """Load the pre-built optimization database"""
        with open(database_file, 'r') as f:
            self.db = json.load(f)
        
        self.profiles = self.db['product_profiles']
        self.transitions = self.db['transition_matrix']
        self.rules = self.db['optimization_rules']
        
        print(f"✅ Loaded database with {len(self.profiles)} products")
    
    def optimize(self, products, wagons_per_product=None):
        """
        Find optimal production sequence
        
        Args:
            products: List of product names, e.g., ['L36', 'L38', 'L30']
            wagons_per_product: Dict of wagons per product (optional)
        
        Returns:
            Optimized sequence and analysis
        """
        if not products:
            return {"error": "No products specified"}
        
        if len(products) == 1:
            return {
                "optimal_sequence": products,
                "total_cost": 0,
                "savings": 0,
                "reason": "Single product - no optimization needed"
            }
        
        # Validate products
        invalid = [p for p in products if p not in self.profiles]
        if invalid:
            return {"error": f"Unknown products: {invalid}"}
        
        # Find optimal sequence
        if len(products) <= 8:
            # Exhaustive search
            best_seq, best_cost = self._exhaustive_search(products)
        else:
            # Intelligent heuristic
            best_seq, best_cost = self._intelligent_sequence(products)
        
        # Calculate savings
        worst_seq, worst_cost = self._worst_case(products)
        savings = ((worst_cost - best_cost) / worst_cost * 100) if worst_cost > 0 else 0
        
        # Generate detailed analysis
        transitions = self._analyze_transitions(best_seq)
        recommendations = self._generate_recommendations(best_seq, wagons_per_product)
        
        return {
            "optimal_sequence": best_seq,
            "total_transition_cost": round(best_cost, 2),
            "worst_case_cost": round(worst_cost, 2),
            "savings_percent": round(savings, 1),
            "transitions": transitions,
            "recommendations": recommendations,
            "estimated_total_energy": self._estimate_total_energy(best_seq, wagons_per_product)
        }
    
    def _exhaustive_search(self, products):
        """Try all permutations"""
        best_seq = None
        best_cost = float('inf')
        
        for perm in permutations(products):
            cost = self._calculate_sequence_cost(perm)
            if cost < best_cost:
                best_cost = cost
                best_seq = list(perm)
        
        return best_seq, best_cost
    
    def _intelligent_sequence(self, products):
        """Smart sequencing for larger sets"""
        # Start with thinnest product
        remaining = set(products)
        thickness_map = {p: self.profiles[p]['thickness_mm'] for p in products}
        
        current = min(remaining, key=lambda p: thickness_map[p])
        sequence = [current]
        remaining.remove(current)
        
        # Build sequence greedily with lookahead
        while remaining:
            best_next = None
            best_score = float('inf')
            
            for next_prod in remaining:
                immediate = self.transitions[current][next_prod]
                
                # Lookahead
                future = 0
                if len(remaining) > 1:
                    temp_remaining = remaining - {next_prod}
                    future = min(self.transitions[next_prod][fp] for fp in temp_remaining) * 0.3
                
                score = immediate + future
                if score < best_score:
                    best_score = score
                    best_next = next_prod
            
            sequence.append(best_next)
            remaining.remove(best_next)
            current = best_next
        
        cost = self._calculate_sequence_cost(sequence)
        return sequence, cost
    
    def _worst_case(self, products):
        """Find worst case scenario"""
        # Alternating thick and thin
        sorted_by_thickness = sorted(products, 
                                     key=lambda p: self.profiles[p]['thickness_mm'])
        
        worst_seq = []
        thin = sorted_by_thickness[:len(products)//2]
        thick = sorted_by_thickness[len(products)//2:][::-1]
        
        for t1, t2 in zip(thin, thick):
            worst_seq.extend([t1, t2])
        
        worst_seq.extend(thin[len(thick):])
        worst_seq.extend(thick[len(thin):])
        
        cost = self._calculate_sequence_cost(worst_seq)
        return worst_seq, cost
    
    def _calculate_sequence_cost(self, sequence):
        """Calculate total transition cost"""
        if len(sequence) < 2:
            return 0
        
        total = 0
        for i in range(len(sequence) - 1):
            total += self.transitions[sequence[i]][sequence[i+1]]
        
        return total
    
    def _analyze_transitions(self, sequence):
        """Detailed transition analysis"""
        transitions = []
        
        for i in range(len(sequence) - 1):
            from_prod = sequence[i]
            to_prod = sequence[i+1]
            
            from_profile = self.profiles[from_prod]
            to_profile = self.profiles[to_prod]
            
            transitions.append({
                "from": from_prod,
                "to": to_prod,
                "cost_kwh": self.transitions[from_prod][to_prod],
                "thickness_change": to_profile['thickness_mm'] - from_profile['thickness_mm'],
                "type_change": from_profile['type'] != to_profile['type'],
                "energy_change": to_profile['avg_kwh_per_m3'] - from_profile['avg_kwh_per_m3']
            })
        
        return transitions
    
    def _generate_recommendations(self, sequence, wagons_per_product):
        """Generate specific recommendations"""
        recommendations = []
        
        # Check for difficult transitions
        for i in range(len(sequence) - 1):
            cost = self.transitions[sequence[i]][sequence[i+1]]
            
            if cost > 100:
                recommendations.append(
                    f"⚠️ High transition cost from {sequence[i]} to {sequence[i+1]} "
                    f"({cost:.1f} kWh). Allow extra setup time."
                )
            
            from_profile = self.profiles[sequence[i]]
            to_profile = self.profiles[sequence[i+1]]
            
            if from_profile['type'] != to_profile['type']:
                recommendations.append(
                    f"🔧 Material type change: {sequence[i]} → {sequence[i+1]}. "
                    f"Schedule cleaning and quality inspection."
                )
        
        # Energy recommendations
        energy_intensive = [
            p for p in sequence 
            if self.profiles[p]['avg_kwh_per_m3'] > 100
        ]
        
        if energy_intensive:
            recommendations.append(
                f"💡 High energy products: {', '.join(energy_intensive)}. "
                f"Consider scheduling during off-peak hours (night)."
            )
        
        # Production volume
        if wagons_per_product:
            total_wagons = sum(wagons_per_product.values())
            if total_wagons > 100:
                recommendations.append(
                    f"📊 High volume week ({total_wagons} wagons). "
                    f"Consider night shifts or split batches."
                )
        
        return recommendations
    
    def _estimate_total_energy(self, sequence, wagons_per_product):
        """Estimate total energy consumption"""
        if not wagons_per_product:
            return None
        
        production_energy = sum(
            self.profiles[prod]['kwh_per_wagon'] * wagons_per_product.get(prod, 0)
            for prod in sequence
        )
        
        transition_energy = sum(
            self.transitions[sequence[i]][sequence[i+1]]
            for i in range(len(sequence) - 1)
        )
        
        return {
            "production_kwh": round(production_energy, 2),
            "transition_kwh": round(transition_energy, 2),
            "total_kwh": round(production_energy + transition_energy, 2)
        }
    
    def get_product_info(self, product):
        """Get detailed info about a product"""
        if product not in self.profiles:
            return {"error": f"Product {product} not found"}
        
        return self.profiles[product]
    
    def compare_sequences(self, seq1, seq2):
        """Compare two different sequences"""
        cost1 = self._calculate_sequence_cost(seq1)
        cost2 = self._calculate_sequence_cost(seq2)
        
        return {
            "sequence_1": seq1,
            "cost_1": cost1,
            "sequence_2": seq2,
            "cost_2": cost2,
            "difference": abs(cost1 - cost2),
            "better": "Sequence 1" if cost1 < cost2 else "Sequence 2",
            "savings_percent": abs(cost1 - cost2) / max(cost1, cost2) * 100
        }


# Example usage
if __name__ == "__main__":
    # Load optimizer
    opt = SimpleProductionOptimizer("optimization_database.json")
    
    # Example weekly demand
    weekly_products = ['L36', 'L38', 'L30', 'L32', 'L40']
    wagons = {
        'L36': 20,
        'L38': 15,
        'L30': 10,
        'L32': 12,
        'L40': 8
    }
    
    # Optimize
    result = opt.optimize(weekly_products, wagons)
    
    # Print results
    print("\n" + "="*60)
    print("PRODUCTION OPTIMIZATION RESULTS")
    print("="*60)
    print(f"\nOptimal Sequence: {' → '.join(result['optimal_sequence'])}")
    print(f"Total Transition Cost: {result['total_transition_cost']} kWh")
    print(f"Savings vs Worst Case: {result['savings_percent']}%")
    
    if result['estimated_total_energy']:
        print(f"\nEstimated Total Energy: {result['estimated_total_energy']['total_kwh']} kWh")
    
    print("\n📋 Transition Details:")
    for trans in result['transitions']:
        print(f"  {trans['from']} → {trans['to']}: {trans['cost_kwh']} kWh")
    
    print("\n💡 Recommendations:")
    for rec in result['recommendations']:
        print(f"  {rec}")
